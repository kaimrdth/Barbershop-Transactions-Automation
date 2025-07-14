/**
 * Combined function that processes data directly from "Raw" to "Processed" 
 * Eliminates the need for the intermediate "Pre-processed" sheet
 */
function processDataComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet names
  const rawSheetName = 'Raw';
  const processedSheetName = 'Processed';
  const commissionSheetName = 'Commission Rates';
  const menuSheetName = 'Menu of Services';
  
  // Get sheets
  const rawSheet = ss.getSheetByName(rawSheetName);
  const processedSheet = ss.getSheetByName(processedSheetName);
  const commissionSheet = ss.getSheetByName(commissionSheetName);
  const menuSheet = ss.getSheetByName(menuSheetName);
  
  // Error handling if any sheet is missing
  if (!rawSheet) throw new Error(`Sheet named '${rawSheetName}' not found.`);
  if (!processedSheet) throw new Error(`Sheet named '${processedSheetName}' not found.`);
  if (!commissionSheet) throw new Error(`Sheet named '${commissionSheetName}' not found.`);
  if (!menuSheet) throw new Error(`Sheet named '${menuSheetName}' not found.`);
  
  // === Step 1: Get and Clean Raw Data ===
  let rawData = rawSheet.getDataRange().getValues();
  
  // Check if there are at least two rows to remove (assuming two header rows in "Raw")
  if (rawData.length < 3) {
    Logger.log("Not enough data in 'Raw' sheet.");
    return;
  }
  
  // Remove the first two rows (headers in Raw)
  rawData.splice(0, 2);
  
  // === Step 2: Define Headers for "Processed" ===
  const headers = [
    'PaymentID',               // A
    'Time & Date',             // B
    'Service Type',            // C
    'Staff Name',              // D
    'Additional Fees',         // E
    'Amount Paid',             // F
    'Processing Fee',          // G
    'Staff Processing Fee',    // H
    'Service Sales',           // I
    'Commission Rate (%)',     // J
    'Staff Service Commission',// K
    'Tips',                    // L
    'Product',                 // M
    'Product Sales',           // N
    'Product Commission Rate', // O
    'Product Commission',      // P
    'Product Tax',             // Q
    'Discounts',               // R
    'Other Adjustments',       // S
    'Total Staff Commission',  // T
    'Net Business Take',       // U
    'Status',                  // V
    'Customer'                 // W
  ];
  
  // === Step 3: Check if "Processed" has headers; if not, set them ===
  let processedLastRow = processedSheet.getLastRow();
  if (processedLastRow === 0) {
    processedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('Headers set in "Processed" sheet.');
    processedLastRow = 1; 
  }
  
  // === Step 4: Get existing PaymentIDs from "Processed" to avoid duplicates ===
  let existingPaymentIDs = new Set();
  if (processedLastRow >= 2) {
    const existingDataRange = processedSheet.getRange(2, 1, processedLastRow - 1, 1); // Col A
    const existingData = existingDataRange.getValues();
    existingData.forEach(row => {
      const paymentID = row[0];
      if (paymentID) {
        existingPaymentIDs.add(paymentID.toString().trim());
      }
    });
  }
  
  // === Step 5: Load Commission and Menu Data ===
  
  // Retrieve commission data from "Commission Rates" (Columns A:C)
  const commissionLastRow = commissionSheet.getLastRow();
  if (commissionLastRow < 2) {
    throw new Error(`No commission rate data found in '${commissionSheetName}'.`);
  }
  const commissionDataRange = commissionSheet.getRange(2, 1, commissionLastRow - 1, 3);
  const commissionData = commissionDataRange.getValues();
  
  // Create a map: { 'Staff Name': { serviceRate, productRate } }
  const commissionMap = {};
  commissionData.forEach(row => {
    const person = row[0].toString().trim();
    const serviceRate = parseFloat(row[1]);
    const productRate = parseFloat(row[2]);
    if (person) {
      commissionMap[person] = {
        serviceRate: isNaN(serviceRate) ? 0 : serviceRate,
        productRate: isNaN(productRate) ? 0 : productRate
      };
    }
  });
  
  // Build a menu map from "Menu of Services": { 'Item Name': Price }
  const menuLastRow = menuSheet.getLastRow();
  const menuDataRange = menuSheet.getRange(2, 1, menuLastRow - 1, 2);
  const menuData = menuDataRange.getValues();
  const menuMap = {};
  menuData.forEach(row => {
    const item = row[0].toString().trim();
    const price = parseFloat(row[1]);
    if (item && !isNaN(price)) {
      menuMap[item] = price;
    }
  });
  
  // Define owners/managers who get different product commission rates
  const owners = ['Owner1', 'Owner2']; // Placeholder names for business owners
  
  // We'll accumulate rows to append to "Processed"
  const processedData = [];
  
  // === Step 6: Process each row from Raw data ===
  rawData.forEach((row, rowIndex) => {
    // Map Raw data fields (accounting for the shifted columns due to new Transaction Date column)
    const date = row[0];                    // Time & Date (A)
    const paymentID = row[1];               // Provider Payment ID (B)
    const currency = row[3];                // Currency (D) - shifted from C
    const amountPaidRaw = row[4];           // Amount (E) - shifted from D
    const processingFeeRaw = row[5];        // Processing Fee (F) - shifted from E
    const commission = parseFloat(row[6]) || 0; // Commission (G) - shifted from F
    const net = row[7];                     // Net (H) - shifted from G
    const statusRaw = row[8];               // Status (I) - shifted from H
    
    // Customer name: concatenate first name (Q) and last name (R) - shifted from P,Q
    const firstName = row[16] || '';
    const lastName = row[17] || '';
    const customerName = (firstName + " " + lastName).trim();
    
    // Order details - all shifted by 1 due to new Transaction Date column
    const orderID = row[44];                // Order ID (AS) - shifted from AR
    const name = row[45];                   // Name (AT) - shifted from AS
    const quantityRaw = row[46];            // Quantity (AU) - shifted from AT
    const discountRaw = row[47];            // Discount (AV) - shifted from AU
    const shipping = row[48];               // Shipping (AW) - shifted from AV
    const productTaxRaw = row[49];          // Tax (AX) - shifted from AW
    
    // Skip rows missing key fields
    if (!name || !date || !paymentID) return;
    
    const trimmedPaymentID = paymentID.toString().trim();
    if (existingPaymentIDs.has(trimmedPaymentID)) {
      // Already processed, skip
      return;
    }
    
    // === Step 7: Parse the "Name" field for staff, service type, products, etc. ===
    const parsedName = parseNameField(name, quantityRaw);
    let staffName = parsedName.staffName;
    let serviceType = parsedName.serviceType;
    const products = parsedName.products;
    const additionalFees = parsedName.additionalFees;
    const nonProductCount = parsedName.nonProductCount;
    
    // Clean up staff/service strings
    staffName = staffName.trim();
    serviceType = serviceType.trim();
    
    // Handle product quantity
    let actualProductQuantity = parseInt(quantityRaw, 10) - nonProductCount;
    if (isNaN(actualProductQuantity) || actualProductQuantity < 1) {
      actualProductQuantity = 1; // fallback
    }
    
    // === Step 8: Calculate Financial Details ===
    
    // Parse numeric fields
    let amountPaid = parseFloat(amountPaidRaw) || 0;
    let processingFee = parseFloat(processingFeeRaw) || 0;
    const halfProcessingFee = processingFee / 2;
    
    // Refund/void check
    const isRefundedOrVoided = (statusRaw === 'Refunded' || statusRaw === 'Voided');
    
    // Assign processing fee: half to staff, half to business, except certain staff
    const noProcessingFeeStaff = ['Product Only'];
    let staffProcessingFee = 0;
    let businessProcessingFee = 0;
    if (noProcessingFeeStaff.includes(staffName)) {
      // Staff doesn't share processing fee
      staffProcessingFee = 0;
      businessProcessingFee = processingFee;
    } else {
      staffProcessingFee = halfProcessingFee;
      businessProcessingFee = halfProcessingFee;
    }
    
    // Look up the service price
    let servicePrice = menuMap[serviceType] || 0;
    
    // Commission rates for staff
    const commissionRates = commissionMap[staffName] || { serviceRate: 0, productRate: 0 };
    const serviceCommissionRate = commissionRates.serviceRate;
    let staffServiceCommission = servicePrice * serviceCommissionRate;
    
    // Calculate product details
    let productSales = 0;
    let productCommissionRate = 0;
    let productCommission = 0;
    let productNames = '';
    
    if (products.length > 0) {
      // Distribute actualProductQuantity among all listed products
      if (products.length === 1) {
        products[0].quantity = actualProductQuantity;
      } else {
        const perProductQuantity = Math.floor(actualProductQuantity / products.length);
        products.forEach(product => {
          product.quantity = perProductQuantity;
        });
        // Any leftover quantity?
        const remaining = actualProductQuantity % products.length;
        for (let i = 0; i < remaining; i++) {
          products[i].quantity += 1;
        }
      }
      
      // Concatenate product names
      productNames = products.map(p => p.name).join(', ');
      
      // Sum up product sales
      products.forEach(product => {
        const productName = product.name.trim();
        const singlePrice = menuMap[productName] || 0;
        productSales += (singlePrice * product.quantity);
      });
      
      // Product commission rate
      if (owners.includes(staffName)) {
        productCommissionRate = 0; 
      } else {
        productCommissionRate = commissionRates.productRate;
      }
      productCommission = productSales * productCommissionRate;
    }
    
    // Additional numeric fields
    let productTax = parseFloat(productTaxRaw) || 0;
    let discounts = parseFloat(discountRaw) || 0;
    
    // Tip calculation (roughly = amountPaid + discounts - services - products - tax)
    let tips = amountPaid + discounts - servicePrice - productSales - productTax;
    if (serviceType === "Product Only") {
      tips = 0;
    }
    
    // Any other adjustments?
    let otherAdjustments = 0; // default
    
    // Total staff commission
    let totalStaffCommission = staffServiceCommission + productCommission - staffProcessingFee + otherAdjustments + tips;
    
    // Net business take
    let netBusinessTake = amountPaid - totalStaffCommission - businessProcessingFee - productTax;
    
    // If refunded/voided, zero out currency fields
    if (isRefundedOrVoided) {
      amountPaid = 0;
      processingFee = 0;
      staffProcessingFee = 0;
      businessProcessingFee = 0;
      servicePrice = 0;
      staffServiceCommission = 0;
      tips = 0;
      productSales = 0;
      productCommission = 0;
      productTax = 0;
      discounts = 0;
      otherAdjustments = 0;
      totalStaffCommission = 0;
      netBusinessTake = 0;
    }
    
    // Round function
    const roundToTwo = num => Math.round(num * 100) / 100;
    
    // === Step 9: Build the final row ===
    processedData.push([
      trimmedPaymentID,                           // A: PaymentID
      date,                                       // B: Time & Date
      serviceType,                                // C: Service Type
      staffName,                                  // D: Staff Name
      additionalFees,                             // E: Additional Fees
      roundToTwo(amountPaid),                     // F: Amount Paid
      roundToTwo(processingFee),                  // G: Processing Fee
      roundToTwo(staffProcessingFee),             // H: Staff Processing Fee
      roundToTwo(servicePrice),                   // I: Service Sales
      commissionRates.serviceRate,                // J: Commission Rate (%)
      roundToTwo(staffServiceCommission),         // K: Staff Service Commission
      roundToTwo(tips),                           // L: Tips
      productNames,                               // M: Product
      roundToTwo(productSales),                   // N: Product Sales
      productCommissionRate,                      // O: Product Commission Rate
      roundToTwo(productCommission),              // P: Product Commission
      roundToTwo(productTax),                     // Q: Product Tax
      roundToTwo(discounts),                      // R: Discounts
      roundToTwo(otherAdjustments),               // S: Other Adjustments
      roundToTwo(totalStaffCommission),           // T: Total Staff Commission
      roundToTwo(netBusinessTake),                // U: Net Business Take
      statusRaw,                                  // V: Status
      customerName                                // W: Customer
    ]);
    
    Logger.log(`Processed PaymentID: ${trimmedPaymentID}, Staff: ${staffName}, Customer: ${customerName}`);
  });
  
  // === Step 10: Append processedData to "Processed" sheet ===
  if (processedData.length > 0) {
    const appendStartRow = processedLastRow + 1;
    processedSheet
      .getRange(appendStartRow, 1, processedData.length, headers.length)
      .setValues(processedData);
    Logger.log(`Appended ${processedData.length} new transactions to "Processed".`);
  } else {
    Logger.log(`No new transactions to process.`);
  }

  // === Step 11: Sort "Processed" sheet by "Time & Date" (column B) descending ===
  const totalRows = processedSheet.getLastRow();
  if (totalRows > 1) {
    processedSheet
      .getRange(2, 1, totalRows - 1, headers.length)
      .sort({column: 2, ascending: false});
    Logger.log(`Sorted "Processed" sheet by "Time & Date" descending.`);
  }

  // === Step 12: Reapply Formatting ===
  applyFormatting(processedSheet, headers.length);
  
  // === Step 13: Remove Duplicate Payment IDs ===
  removeDuplicatePaymentIDs();
  
  Logger.log(`Processing complete! Processed ${processedData.length} new transactions from Raw to Processed.`);
}

/**
 * Removes duplicate rows in the "Processed" sheet based on "PaymentID" (column A).
 */
function removeDuplicatePaymentIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const processedSheet = ss.getSheetByName('Processed');
  if (!processedSheet) throw new Error(`Sheet named 'Processed' not found.`);
  
  const lastRow = processedSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log(`No data to process in 'Processed' sheet.`);
    return;
  }
  
  const dataRange = processedSheet.getRange(2, 1, lastRow - 1, processedSheet.getLastColumn());
  const data = dataRange.getValues();
  const seenPaymentIDs = new Set();
  const rowsToDelete = [];
  
  // Iterate from bottom up so we delete duplicates after the first occurrence
  for (let i = data.length - 1; i >= 0; i--) {
    const paymentID = data[i][0]; // Column A
    if (paymentID) {
      const trimmedPaymentID = paymentID.toString().trim();
      if (seenPaymentIDs.has(trimmedPaymentID)) {
        rowsToDelete.push(i + 2); 
      } else {
        seenPaymentIDs.add(trimmedPaymentID);
      }
    }
  }
  
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => a - b);
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      processedSheet.deleteRow(rowsToDelete[i]);
    }
    Logger.log(`Removed ${rowsToDelete.length} duplicate row(s) based on "PaymentID".`);
  } else {
    Logger.log(`No duplicate "PaymentID" entries found.`);
  }
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

  Logger.log('Reapplied formatting to the "Processed" sheet.');
}

/**
 * Parses the "Name" field from Raw to extract staff name, service type, product info, etc.
 * @param {string} name - The concatenated name string from raw data
 * @param {number|string} quantityRaw - The raw quantity field
 * @returns {object} { staffName, serviceType, products[], additionalFees, nonProductCount }
 */
function parseNameField(name, quantityRaw) {
  let staffName = '';
  let serviceType = '';
  let products = [];
  let additionalFees = 'No';
  let nonProductCount = 0;
  
  let quantity = parseInt(quantityRaw, 10);
  if (isNaN(quantity) || quantity < 1) {
    quantity = 1; 
  }

  // Split by commas
  const parts = name.split(',').map(part => part.trim());
  
  parts.forEach(part => {
    // Regex to match 'Service Type w/ Staff Name'
    const serviceStaffRegex = /^(.*?)\s*w\/\s*(.*)$/i;
    const match = part.match(serviceStaffRegex);
    
    if (match) {
      serviceType = match[1].trim();
      staffName = match[2].trim();
      nonProductCount++; // one service item
    } else if (part.toLowerCase() === 'additional fees') {
      additionalFees = 'Yes';
      nonProductCount++; 
    } else if (part !== '') {
      // Assume it's a product
      products.push({ name: part, quantity: 1 });
    }
  });
  
  // If only products are present, treat as "Product Only"
  if (!staffName && !serviceType && products.length > 0) {
    staffName = "Product Only";
    serviceType = "Product Only";
  }
  
  return {
    staffName: staffName,
    serviceType: serviceType,
    products: products,
    additionalFees: additionalFees,
    nonProductCount: nonProductCount
  };
}