/*******************************
 * Kinship Barbershop — Square → Google Sheet
 * Destination tab: Processed
 * Requires: Script Property SQUARE_ACCESS_TOKEN
 * Commission table: sheet "Commission Rates"
 *   A: Person
 *   B: Service Commission Rate
 *   C: Product Commission Rate
 *   D: Square Team Member ID
 *******************************/

const SQUARE_API_BASE = 'https://connect.squareup.com/v2';
const SQUARE_VERSION = '2025-07-16'; // Update as Square releases new versions
const DEST_TAB = 'Processed';
const COMMISSION_SHEET_NAME = 'Commission Rates';

const SYNC_CURSOR_KEY = 'SQUARE_UPDATED_CURSOR_ISO';
const TEAM_CACHE_KEY = 'SQUARE_TEAM_CACHE_JSON'; // team_member_id -> "Given Family"
const CUSTOMER_CACHE_KEY = 'SQUARE_CUSTOMER_CACHE_JSON'; // customer_id -> "Given Family"
const BOOKING_CACHE_KEY = 'SQUARE_BOOKING_STAFF_CACHE_JSON'; // appointment_id -> team_member_id
const DEFAULT_LOOKBACK_DAYS = 30;

// Toggle: enable diagnostic logging for missing staff scenarios
const ENABLE_MISSING_STAFF_LOGS = true;

// Defaults only. The sheet overrides these per staff.
const COMMISSION_RULES = {
  defaultServiceRate: 0.0,
  defaultProductRate: 0.0,
  byItemName: {
    // Optional. Example:
    // 'Beard Trim': { service: 0.4 },
  }
};

const HEADERS = [
  'PaymentID','Time & Date','Service Type','Staff Name','Additional Fees',
  'Amount Paid','Processing Fee','Staff Processing Fee','Service Sales',
  'Commission Rate (%)','Staff Service Commission','Tips','Product','Product Sales',
  'Product Commission Rate','Product Commission','Product Tax','Discounts',
  'Other Adjustments','Total Staff Commission','Net Business Take','Status','Customer','Flags'
];

function syncSquareToSheet() {
  const sheet = getOrCreateSheet_(DEST_TAB, HEADERS);
  const props = PropertiesService.getScriptProperties();

  const nowIso = new Date().toISOString();
  const beginIso = props.getProperty(SYNC_CURSOR_KEY) || isoDaysAgo_(DEFAULT_LOOKBACK_DAYS);

  const payments = fetchPaymentsUpdatedSince_(beginIso, nowIso);
  if (!payments.length) {
    Logger.log('No new or updated payments.');
    props.setProperty(SYNC_CURSOR_KEY, nowIso);
    return;
  }

  const orderIds = unique_(payments.map(p => p.order_id).filter(Boolean));
  const teamIdsFromPayments = unique_(payments.map(p => p.team_member_id).filter(Boolean));

  const ordersById = orderIds.length ? batchRetrieveOrders_(orderIds) : {};
  const variationIds = collectLineVariationIds_(ordersById);
  const catalogInfo = variationIds.length ? batchRetrieveCatalogMap_(variationIds) : initCatalogInfo_();
  
  // Read commission table with team member IDs
  const commissionData = readCommissionRatesWithTeamIds_();

  // Prefetch booking -> staff for orders tied to appointments
  const bookingStaffByApptId = prefetchBookingStaff_(ordersById);

  // Combine team member IDs from payments and bookings
  const teamIdsFromBookings = unique_(Object.values(bookingStaffByApptId).filter(Boolean));
  const teamIds = unique_([...(teamIdsFromPayments || []), ...teamIdsFromBookings]);

  // Build enhanced staff lookup
  const staffById = buildStaffLookup_(teamIds, commissionData.teamIdToName);

  const orderCustomerIds = unique_(Object.values(ordersById)
    .map(o => o && o.customer_id)
    .filter(Boolean));

  const paymentCustomerIds = unique_(payments
    .map(p => p && p.customer_id)
    .filter(Boolean));

  const allCustomerIds = unique_([...paymentCustomerIds, ...orderCustomerIds]);
  const customersById = allCustomerIds.length ? bulkRetrieveCustomers_(allCustomerIds) : {};

  const paymentRowIndex = buildExistingIndex_(sheet, 1); // PaymentID -> row

  const updates = [];
  const appends = [];
  payments.forEach(p => {
    const row = buildProcessedRow_(p, ordersById[p.order_id], catalogInfo, staffById, customersById, commissionData.commissionByPerson, bookingStaffByApptId);
    const paymentId = String(row[0]);
    if (paymentRowIndex.hasOwnProperty(paymentId)) {
      updates.push({row: paymentRowIndex[paymentId], values: row});
    } else {
      appends.push(row);
    }
  });

  if (updates.length) {
    const lastCol = colLetter_(HEADERS.length);
    const range = sheet.getRangeList(updates.map(u => `A${u.row}:${lastCol}${u.row}`)).getRanges();
    updates.forEach((u, i) => range[i].setValues([u.values]));
  }
  if (appends.length) {
    sheet.getRange(sheet.getLastRow()+1, 1, appends.length, HEADERS.length).setValues(appends);
  }

  // Apply formatting after updates/appends
  applyFormatting(sheet, HEADERS.length);

  // Sort by "Time & Date" (column B) descending
  const totalRows = sheet.getLastRow();
  if (totalRows > 1) {
    sheet
      .getRange(2, 1, totalRows - 1, HEADERS.length)
      .sort({column: 2, ascending: false});
    Logger.log(`Sorted "Processed" sheet by "Time & Date" descending.`);
  }

  props.setProperty(SYNC_CURSOR_KEY, nowIso);
  Logger.log(`Processed ${payments.length} payments. Updated ${updates.length}, appended ${appends.length}.`);
}

function buildProcessedRow_(payment, order, catalogInfo, staffById, customersById, commissionByPerson, bookingStaffByApptId) {
  const money = m => (m && typeof m.amount === 'number') ? m.amount / 100 : 0;
  const paymentId = payment.id;
  const createdAtIso = payment.created_at || payment.updated_at || '';
  const createdAt = createdAtIso ? formatDateTime_(createdAtIso) : '';
  const status = payment.status || '';
  const tips = money(payment.tip_money);
  const amountPaid = money(payment.total_money);
  const refundedTotal = money(payment.refunded_money);
  const processingFee = (payment.processing_fee || [])
    .map(f => money(f.amount_money || f.applied_money))
    .reduce((a,b)=>a+b, 0);

  // Prefer booking staff, then payment, then legacy order field
  let fullStaffName = '';
  let staffFlag = '';
  const apptId = extractAppointmentIdFromOrder_(order);
  if (apptId && bookingStaffByApptId && bookingStaffByApptId[apptId]) {
    const tmId = bookingStaffByApptId[apptId];
    fullStaffName = staffById[tmId] || '';
    staffFlag = 'from_booking';
  } else if (payment.team_member_id) {
    fullStaffName = staffById[payment.team_member_id] || '';
    staffFlag = 'from_payment';
  } else if (order && order.employee_id) {
    fullStaffName = staffById[order.employee_id] || '';
    staffFlag = 'from_order_legacy';
  } else {
    staffFlag = 'STAFF_MISSING';
  }
  const staffName = fullStaffName ? fullStaffName.split(' ')[0] : '';
  if (staffFlag === 'STAFF_MISSING' && ENABLE_MISSING_STAFF_LOGS) {
    logMissingStaffDiagnostic_(payment, order, bookingStaffByApptId);
  }
  const chosenCustomerId = payment.customer_id || (order && order.customer_id) || null;
  const customerName = resolveCustomerName_(chosenCustomerId, customersById, payment);

  let serviceNames = [];
  let productNames = [];
  let serviceSales = 0;
  let productSales = 0;
  let productTax = 0;
  let discounts = 0;
  let additionalFees = 0;

  if (order) {
    const lineItems = (order.line_items || []);
    lineItems.forEach(li => {
      const varId = li.catalog_object_id;
      const parentInfo = catalogInfo.variationToItem[varId] || {};
      const isService = parentInfo.product_type === 'APPOINTMENTS_SERVICE';
      const gross = money(li.gross_sales_money);
      const lineDiscount = money(li.total_discount_money);
      const net = Math.max(0, gross - lineDiscount);
      const name = (parentInfo.item_name || li.name || '').trim();

      if (isService) {
        serviceNames.push(name);
        serviceSales += net;
      } else {
        productNames.push(name);
        productSales += net;
        productTax += money(li.total_tax_money);
      }
    });

    discounts = money(order.total_discount_money);
    additionalFees = money(order.total_service_charge_money);
    if (!additionalFees && order.service_charges) {
      additionalFees = order.service_charges
        .map(sc => money(sc.applied_money || sc.total_money))
        .reduce((a,b)=>a+b, 0);
    }
  }

  const serviceLabel = unique_(serviceNames.filter(Boolean)).join(', ');
  const productLabel = unique_(productNames.filter(Boolean)).join(', ');

  // Commission resolution order:
  // 1) Item-specific override from COMMISSION_RULES.byItemName
  // 2) Staff rate from Commission Rates sheet
  // 3) Default from COMMISSION_RULES
  const svcRate = resolveCommissionRate_(staffName, serviceLabel, true, commissionByPerson);
  const prodRate = resolveCommissionRate_(staffName, productLabel, false, commissionByPerson);

  const staffServiceCommission = round2_(serviceSales * svcRate);
  const productCommission     = round2_(productSales * prodRate);
  const staffProcessingFee = 0; // set if you share fees with staff
  const totalStaffCommission = round2_(staffServiceCommission + productCommission + tips - staffProcessingFee);

  const otherAdjustments = 0;
  const netBusinessTake = round2_(
    amountPaid
    - processingFee
    - totalStaffCommission
    - tips
    - refundedTotal
    + additionalFees
    - discounts
    + otherAdjustments
  );

  return [
    paymentId,
    createdAt,
    serviceLabel,
    staffName,
    toFixedOrBlank_(additionalFees),

    toFixedOrBlank_(amountPaid),
    toFixedOrBlank_(processingFee),
    toFixedOrBlank_(staffProcessingFee),
    toFixedOrBlank_(serviceSales),

    svcRate || 0,
    toFixedOrBlank_(staffServiceCommission),
    toFixedOrBlank_(tips),
    productLabel,
    toFixedOrBlank_(productSales),

    prodRate || 0,
    toFixedOrBlank_(productCommission),
    toFixedOrBlank_(productTax),
    toFixedOrBlank_(discounts),

    toFixedOrBlank_(otherAdjustments),
    toFixedOrBlank_(totalStaffCommission),
    toFixedOrBlank_(netBusinessTake),
    status,
    customerName,
    staffFlag
  ];
}
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function extractAppointmentIdFromOrder_(order) {
  if (!order) return null;
  const fulf = order.fulfillments || [];
  for (const f of fulf) {
    const det = f.appointment_details || f.metadata || {};
    if (det && det.appointment_id) return det.appointment_id;
    if (det && det.booking_id) return det.booking_id;
  }
  return null;
}

function truncateJson_(obj, maxChars) {
  try {
    const s = JSON.stringify(obj);
    if (s.length <= maxChars) return s;
    return s.slice(0, Math.max(0, maxChars - 20)) + '... [truncated]';
  } catch (e) {
    return String(obj);
  }
}

/**
 * Logs a structured diagnostic payload when staff cannot be resolved.
 * Includes raw payment, order, and booking (if available). Keeps under Apps Script log limits.
 */
function logMissingStaffDiagnostic_(payment, order, bookingStaffByApptId) {
  try {
    const apptId = extractAppointmentIdFromOrder_(order);
    let bookingRaw = null;
    let bookingSegments = null;
    if (apptId) {
      try {
        const br = squareGet_(`/bookings/${encodeURIComponent(apptId)}`);
        bookingRaw = br;
        const b = br && (br.booking || br);
        bookingSegments = b && b.appointment_segments ? b.appointment_segments : null;
      } catch (e) {
        bookingRaw = { error: String(e) };
      }
    }

    const diag = {
      tag: 'STAFF_MISSING',
      payment_id: payment && payment.id,
      order_id: order && order.id,
      appt_id: apptId || null,
      booking_team_member_id: apptId ? ((bookingStaffByApptId && bookingStaffByApptId[apptId]) || null) : null,
      payment_team_member_id: (payment && payment.team_member_id) || null,
      order_employee_id: (order && order.employee_id) || null,
      location_id: payment && payment.location_id,
      created_at: payment && payment.created_at,
      updated_at: payment && payment.updated_at,
      order_customer_id: order && order.customer_id,
      payment_customer_id: payment && payment.customer_id,
      order_note: order && order.note,
      line_item_names: order ? ((order.line_items || []).map(li => li && li.name).filter(Boolean)) : [],
      order_fulfillments: order && order.fulfillments,
      booking_segments: bookingSegments,
      payment_raw: payment,
      order_raw: order,
      booking_raw: bookingRaw
    };

    // Keep logs within ~130KB to avoid truncation (Apps Script limit ~256KB per execution)
    Logger.log(truncateJson_(diag, 130000));
  } catch (e) {
    Logger.log('logMissingStaffDiagnostic_ error: ' + e);
  }
}

/**
 * Prefetch booking -> team_member_id for all orders that have an appointment.
 * Caches results in Script Properties.
 * Returns: { appointment_id: team_member_id, ... }
 */
function prefetchBookingStaff_(ordersById) {
  const props = PropertiesService.getScriptProperties();
  const cache = JSON.parse(props.getProperty(BOOKING_CACHE_KEY) || '{}');
  const apptIds = new Set();

  Object.values(ordersById).forEach(o => {
    const apptId = extractAppointmentIdFromOrder_(o);
    if (apptId) apptIds.add(apptId);
  });

  const missing = Array.from(apptIds).filter(id => !(id in cache));
  missing.forEach(id => {
    try {
      const res = squareGet_(`/bookings/${encodeURIComponent(id)}`);
      const booking = res.booking || res;
      const segs = booking && booking.appointment_segments ? booking.appointment_segments : [];
      const tmId = (segs && segs.length) ? segs[0].team_member_id : '';
      cache[id] = tmId || '';
      Utilities.sleep(120);
    } catch (e) {
      Logger.log(`Booking fetch failed for ${id}: ${e}`);
      cache[id] = '';
    }
  });

  props.setProperty(BOOKING_CACHE_KEY, JSON.stringify(cache));

  const out = {};
  Array.from(apptIds).forEach(id => out[id] = cache[id] || '');
  return out;
}

/**
 * Read "Commission Rates" sheet including Square Team Member IDs from column D
 * Returns: 
 * - commissionByPerson: { 'First Last': {service: 0.x, product: 0.y}, ... }
 * - teamIdToName: { 'team_member_id': 'First Last', ... }
 */
function readCommissionRatesWithTeamIds_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(COMMISSION_SHEET_NAME);
  const commissionByPerson = {};
  const teamIdToName = {};
  
  if (!sh) return { commissionByPerson, teamIdToName };

  const last = sh.getLastRow();
  if (last < 2) return { commissionByPerson, teamIdToName };

  const values = sh.getRange(2, 1, last - 1, 4).getValues(); // A:D
  values.forEach(r => {
    const name = String(r[0] || '').trim();
    if (!name) return;
    
    const svc = normalizeRate_(r[1]);
    const prod = normalizeRate_(r[2]);
    const teamId = String(r[3] || '').trim();
    
    commissionByPerson[name] = { service: svc, product: prod };
    
    // Map team member ID to person name
    if (teamId) {
      teamIdToName[teamId] = name;
      Logger.log(`Mapped team member ${teamId} to ${name}`);
    }
  });
  
  return { commissionByPerson, teamIdToName };
}

/**
 * Build staff lookup with fallback to Commission Rates mapping
 */
function buildStaffLookup_(teamIds, teamIdToName) {
  const props = PropertiesService.getScriptProperties();
  const cache = JSON.parse(props.getProperty(TEAM_CACHE_KEY) || '{}');
  const out = {};
  const missing = [];
  
  teamIds.forEach(id => {
    // First check if we have a mapping from Commission Rates
    if (teamIdToName[id]) {
      out[id] = teamIdToName[id];
      // Update cache with this mapping
      cache[id] = teamIdToName[id];
    } else if (cache[id]) {
      out[id] = cache[id];
    } else {
      missing.push(id);
    }
  });
  
  // For any still missing, try to fetch from Square API
  missing.forEach(id => {
    try {
      const res = squareGet_(`/team-members/${encodeURIComponent(id)}`);
      const tm = res.team_member || {};
      
      // Try multiple fields for the name
      let name = '';
      
      if (tm.display_name && tm.display_name.trim()) {
        name = tm.display_name.trim();
      } else if (tm.given_name && tm.given_name.trim()) {
        const parts = [tm.given_name.trim()];
        if (tm.family_name && tm.family_name.trim()) {
          parts.push(tm.family_name.trim());
        }
        name = parts.join(' ');
      } else if (tm.family_name && tm.family_name.trim()) {
        name = tm.family_name.trim();
      }
      
      if (name) {
        cache[id] = name;
        out[id] = name;
        Logger.log(`Fetched team member ${id}: ${name}`);
      } else {
        Logger.log(`Warning: No name found for team member ${id}`);
      }
      
      Utilities.sleep(120);
    } catch (e) {
      Logger.log(`Team fetch failed for ${id}: ${e}`);
    }
  });
  
  props.setProperty(TEAM_CACHE_KEY, JSON.stringify(cache));
  return out;
}

/** Accepts numbers like 0.5, 0, 50, "50%", "" and returns decimal rate. Blank -> 0. */
function normalizeRate_(val) {
  if (val === '' || val === null || val === undefined) return 0;
  if (typeof val === 'number') {
    return val > 1 ? val / 100 : val;
  }
  const s = String(val).trim();
  if (!s) return 0;
  if (s.endsWith('%')) {
    const n = parseFloat(s.slice(0, -1));
    return isFinite(n) ? n / 100 : 0;
  }
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return n > 1 ? n / 100 : n;
}

/** Decide rate using item override, then sheet, then defaults. */
function resolveCommissionRate_(staffName, itemLabel, isService, commissionByPerson) {
  // 1) Item override
  const byItem = COMMISSION_RULES.byItemName || {};
  for (const key in byItem) {
    if (key && itemLabel && itemLabel.includes(key)) {
      const v = byItem[key];
      if (isService && typeof v.service === 'number') return v.service;
      if (!isService && typeof v.product === 'number') return v.product;
    }
  }
  // 2) Person from sheet
  const rec = commissionByPerson[staffName];
  if (rec) return isService ? (rec.service ?? 0) : (rec.product ?? 0);

  // 3) Defaults
  return isService ? COMMISSION_RULES.defaultServiceRate : COMMISSION_RULES.defaultProductRate;
}

// ===== Square fetchers =====

function fetchPaymentsUpdatedSince_(updatedBeginIso, updatedEndIso) {
  const results = [];
  let cursor = null;
  do {
    const params = {
      updated_at_begin_time: updatedBeginIso,
      updated_at_end_time: updatedEndIso,
      sort_order: 'ASC',
      limit: 100
    };
    if (cursor) params.cursor = cursor;
    const res = squareGet_('/payments', params);
    if (res && res.payments) results.push(...res.payments);
    cursor = res && res.cursor;
  } while (cursor);
  return results;
}

function batchRetrieveOrders_(orderIds) {
  const map = {};
  for (let i = 0; i < orderIds.length; i += 100) {
    const chunk = orderIds.slice(i, i+100);
    const res = squarePost_('/orders/batch-retrieve', { order_ids: chunk });
    (res.orders || []).forEach(o => { map[o.id] = o; });
    Utilities.sleep(150);
  }
  return map;
}

function collectLineVariationIds_(ordersById) {
  const ids = new Set();
  Object.values(ordersById).forEach(o => {
    (o.line_items || []).forEach(li => {
      if (li.catalog_object_id) ids.add(li.catalog_object_id);
    });
  });
  return Array.from(ids);
}

function initCatalogInfo_() {
  return { variationToItem: {} };
}

function batchRetrieveCatalogMap_(variationIds) {
  const variationToItem = {};
  for (let i = 0; i < variationIds.length; i += 100) {
    const chunk = variationIds.slice(i, i+100);
    const res = squarePost_('/catalog/batch-retrieve', {
      object_ids: chunk,
      include_related_objects: true
    });
    const objects = (res.objects || []);
    const related = (res.related_objects || []);
    const all = [...objects, ...related];

    const itemsById = {};
    all.forEach(obj => {
      if (obj.type === 'ITEM' && obj.item_data) {
        itemsById[obj.id] = {
          product_type: obj.item_data.product_type,
          item_name: obj.item_data.name
        };
      }
    });

    all.forEach(obj => {
      if (obj.type === 'ITEM_VARIATION' && obj.item_variation_data) {
        const parentId = obj.item_variation_data.item_id;
        const parent = itemsById[parentId] || {};
        variationToItem[obj.id] = {
          product_type: parent.product_type,
          item_name: parent.item_name
        };
      }
    });
    Utilities.sleep(150);
  }
  return { variationToItem };
}

function bulkRetrieveCustomers_(customerIds) {
  const props = PropertiesService.getScriptProperties();
  const cache = JSON.parse(props.getProperty(CUSTOMER_CACHE_KEY) || '{}');
  const out = {};
  const missing = customerIds.filter(id => !cache[id]);
  for (let i = 0; i < missing.length; i += 100) {
    const chunk = missing.slice(i, i+100);
    try {
      const res = squarePost_('/customers/bulk-retrieve', { customer_ids: chunk });
      const mapCandidate = res.customers || res.responses || {};
      Object.keys(mapCandidate).forEach(key => {
        const entry = mapCandidate[key];
        const c = (entry && entry.customer) ? entry.customer : entry;
        if (!c) return;
        const cid = c.id || key; // prefer object id when present
        const name = [c.given_name, c.family_name].filter(Boolean).join(' ').trim()
          || c.company_name || c.email_address || '';
        if (cid && name) cache[cid] = name;
      });
      Utilities.sleep(150);
    } catch (e) {
      Logger.log(`Customers bulk-retrieve failed: ${e}`);
    }
  }
  customerIds.forEach(id => out[id] = cache[id] || '');
  props.setProperty(CUSTOMER_CACHE_KEY, JSON.stringify(cache));
  return out;
}

// ===== HTTP helpers =====

function squareHeaders_() {
  const token = PropertiesService.getScriptProperties().getProperty('SQUARE_ACCESS_TOKEN');
  if (!token) throw new Error('Missing SQUARE_ACCESS_TOKEN in Script Properties.');
  return {
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json',
    'Square-Version': SQUARE_VERSION
  };
}

function squareGet_(path, params) {
  let url = `${SQUARE_API_BASE}${path}`;
  if (params) {
    const qs = Object.keys(params)
      .filter(k => params[k] !== undefined && params[k] !== null && params[k] !== '')
      .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])))
      .join('&');
    if (qs) url += `?${qs}`;
  }
  const resp = UrlFetchApp.fetch(url, { method: 'get', headers: squareHeaders_(), muteHttpExceptions: true });
  return parseResponse_(resp, url, 'GET');
}

function squarePost_(path, body) {
  const url = `${SQUARE_API_BASE}${path}`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: squareHeaders_(),
    contentType: 'application/json',
    payload: JSON.stringify(body || {}),
    muteHttpExceptions: true
  });
  return parseResponse_(resp, url, 'POST');
}

function parseResponse_(resp, url, method) {
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code >= 200 && code < 300) {
    return text ? JSON.parse(text) : {};
  }
  throw new Error(`${method} ${url} failed (${code}): ${text}`);
}

// ===== Sheet helpers =====

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = firstRow.some((v, i) => String(v).trim() !== headers[i]);
  if (needsHeaders) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function buildExistingIndex_(sheet, headerRows) {
  const last = sheet.getLastRow();
  const index = {};
  if (last <= headerRows) return index;
  const range = sheet.getRange(headerRows+1, 1, last-headerRows, 1).getValues(); // col A
  range.forEach((row, i) => {
    const id = String(row[0] || '').trim();
    if (id) index[id] = headerRows + 1 + i;
  });
  return index;
}

// ===== Generic helpers =====

function isoDaysAgo_(d) {
  const dt = new Date();
  dt.setDate(dt.getDate() - d);
  return dt.toISOString();
}
function unique_(arr) { return Array.from(new Set(arr)); }
function round2_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
function toFixedOrBlank_(n) { return (typeof n === 'number' && isFinite(n)) ? round2_(n) : ''; }

function formatDateTime_(isoString) {
  try {
    const date = new Date(isoString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const hours = date.getHours();
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const seconds = date.getSeconds().toString().padStart(2, '0');
    return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
  } catch (e) {
    return isoString; // fallback to original if parsing fails
  }
}

function resolveCustomerName_(chosenCustomerId, customersById, payment) {
  if (chosenCustomerId && customersById[chosenCustomerId]) return customersById[chosenCustomerId];
  // fallbacks when no customer profile is linked
  const addr = payment.billing_address || payment.shipping_address || {};
  const addrName = [addr.first_name, addr.last_name].filter(Boolean).join(' ').trim()
    || [addr.given_name, addr.family_name].filter(Boolean).join(' ').trim();
  if (addrName) return addrName;
  const cardName = payment.card_details && payment.card_details.card && payment.card_details.card.cardholder_name;
  if (cardName) return cardName;
  return payment.buyer_email_address || '';
}

/**
 * Applies formatting to the "Processed" sheet (background colors, bold headers, auto-resize, etc.).
 * @param {Sheet} sheet 
 * @param {number} headerLength 
 */
function applyFormatting(sheet, headerLength) {
  const lastRow = sheet.getLastRow();
  
  // Background colors for columns G(7) & H(8): light blue
  const columnsGH = [7, 8];
  columnsGH.forEach(function(col){
    const range = sheet.getRange(1, col, lastRow);
    range.setBackground('#D9E1F2'); // Light blue
  });

  // Columns I(9), J(10), K(11), L(12): light green
  const columnsIJKL = [9, 10, 11, 12];
  columnsIJKL.forEach(function(col){
    const range = sheet.getRange(1, col, lastRow);
    range.setBackground('#E2EFDA'); // Light green
  });

  // Columns M(13), N(14), O(15), P(16), Q(17): light yellow
  const columnsMNOPQ = [13, 14, 15, 16, 17];
  columnsMNOPQ.forEach(function(col){
    const range = sheet.getRange(1, col, lastRow);
    range.setBackground('#FFF2CC'); // Light yellow
  });

  // Columns R(18) and S(19): light pink
  const columnsRS = [18, 19];
  columnsRS.forEach(function(col){
    const range = sheet.getRange(1, col, lastRow);
    range.setBackground('#FCE4EC'); // Light pink
  });

  // Bold headers
  sheet.getRange(1, 1, 1, headerLength).setFontWeight('bold');

  // Auto-resize columns
  sheet.autoResizeColumns(1, headerLength);

  // For rows after the header
  if (lastRow > 1) {
    // Format currency columns
    const currencyColumns = [6,7,8,9,11,12,14,16,17,18,19,20,21]; 
    currencyColumns.forEach(function(col){
      const range = sheet.getRange(2, col, lastRow - 1);
      range.setNumberFormat('$#,##0.00');
    });

    // Format percent columns
    const percentColumns = [10,15];
    percentColumns.forEach(function(col){
      const range = sheet.getRange(2, col, lastRow - 1);
      range.setNumberFormat('0.00%');
    });
  }

  Logger.log('Applied formatting to the "Processed" sheet.');
}

// Force refresh - clears all cursors and caches for a full re-sync
function forceRefresh() {
  const props = PropertiesService.getScriptProperties();
  
  // Clear sync cursor to force re-processing from default lookback
  props.deleteProperty(SYNC_CURSOR_KEY);
  
  // Clear caches to force fresh API calls
  props.deleteProperty(TEAM_CACHE_KEY);
  props.deleteProperty(CUSTOMER_CACHE_KEY);
  props.deleteProperty(BOOKING_CACHE_KEY);
  
  Logger.log('Force refresh: Cleared sync cursor and caches. Next sync will re-process last 30 days.');
  
  // Optionally run sync immediately
  // syncSquareToSheet();
}

// Optional trigger
function createHourlyTrigger() {
  ScriptApp.newTrigger('syncSquareToSheet').timeBased().everyHours(1).create();
}

/**
 * Diagnostic function to verify Commission Rates mapping
 */
function verifyCommissionRatesSetup() {
  const data = readCommissionRatesWithTeamIds_();
  
  Logger.log('Commission Rates Setup:');
  Logger.log('========================');
  
  // Show person -> commission mapping
  Object.entries(data.commissionByPerson).forEach(([name, rates]) => {
    Logger.log(`${name}: Service ${rates.service * 100}%, Product ${rates.product * 100}%`);
  });
  
  Logger.log('\nTeam Member ID Mappings:');
  Logger.log('========================');
  
  // Show team ID -> name mapping
  if (Object.keys(data.teamIdToName).length === 0) {
    Logger.log('No team member IDs found in column D. Please add Square Team Member IDs.');
  } else {
    Object.entries(data.teamIdToName).forEach(([teamId, name]) => {
      Logger.log(`${teamId} → ${name}`);
    });
  }
  
  return data;
}