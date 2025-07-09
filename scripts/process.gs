/**
 * Processes transactions from "Raw Transactions" and appends new entries to the "Processed Transactions" sheet.
 * Handles commission calculations, product sales, and business metrics.
 */
function processBusinessTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet names
  const rawSheetName = 'Raw Transactions';
  const processedSheetName = 'Processed Transactions';
  const commissionSheetName = 'Commission Rates';
  const menuSheetName = 'Service Menu';
  
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
  
  // Define headers for "Processed Transactions"
  const headers = [
    'Transaction ID',          // A
    'Date & Time',             // B
    'Service Type',            // C
    'Staff Member',            // D
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
    'Net Business Revenue',    // U
    'Status',                  // V
    'Customer Name'            // W
  ];
  
  // Check if "Processed Transactions" has headers; if not, set them
  let processedLastRow = processedSheet.getLastRow();
  if (processedLastRow === 0) {
    processedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('Headers set in "Processed Transactions" sheet.');
    processedLastRow = 1; 
  }
  
  // Get existing Transaction IDs from "Processed Transactions" to avoid duplicates
  let existingTransactionIDs = new Set();
  if (processedLastRow >= 2) {
    const existingDataRange = processedSheet.getRange(2, 1, processedLastRow - 1, 1); // Col A
    const existingData = existingDataRange.getValues();
    existingData.forEach(row => {
      const transactionID = row[0];
      if (transactionID) {
        existingTransactionIDs.add(transactionID.toString().trim());
      }
    });
  }
  
  // Get raw transaction data, starting from row 2
  const rawLastRow = rawSheet.getLastRow();
  const rawLastColumn = rawSheet.getLastColumn();
  if (rawLastRow < 2) {
    Logger.log(`No data found in '${rawSheetName}' starting from row 2.`);
    return;
  }
  const rawDataRange = rawSheet.getRange(2, 1, rawLastRow - 1, rawLastColumn);
  const rawData = rawDataRange.getValues();
  
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
  
  // Build a menu map from "Service Menu": { 'Item Name': Price }
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
  
  // Define business owners, who get different product commission rates
  const businessOwners = ['Manager A', 'Manager B'];
  
  // We'll accumulate rows to append to "Processed Transactions"
  const processedData = [];
  
  // === Process each row from "Raw Transactions" ===
  rawData.forEach((row, rowIndex) => {
    // Map raw data columns to variables
    const date = row[0];                // col A
    const transactionID = row[1];       // col B
    const amountPaidRaw = row[3];       // col D
    const processingFeeRaw = row[5];    // col F
    const statusRaw = row[9];           // col J
    const customerName = row[10];       // col K
    const serviceDescription = row[12]; // col M -> "Service Description"
    const quantityRaw = row[13];        // col N -> "Quantity"
    const discountRaw = row[14];        // col O -> "Discount"
    const productTaxRaw = row[16];      // col Q -> "Tax"
    
    // Skip rows missing key fields
    if (!serviceDescription || !date || !transactionID) return;
    
    const trimmedTransactionID = transactionID.toString().trim();
    if (existingTransactionIDs.has(trimmedTransactionID)) {
      // Already processed, skip
      return;
    }
    
    // Parse the service description field for staff, service type, products, etc.
    const parsedService = parseServiceDescription(serviceDescription, quantityRaw);
    let staffName = parsedService.staffName;
    let serviceType = parsedService.serviceType;
    const products = parsedService.products;
    const additionalFees = parsedService.additionalFees;
    const nonProductCount = parsedService.nonProductCount;
    
    // Clean up staff/service strings
    staffName = staffName.trim();
    serviceType = serviceType.trim();
    
    // Handle product quantity
    let actualProductQuantity = parseInt(quantityRaw, 10) - nonProductCount;
    if (isNaN(actualProductQuantity) || actualProductQuantity < 1) {
      actualProductQuantity = 1; // fallback
    }
    
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
      if (businessOwners.includes(staffName)) {
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
    
    // Net business revenue
    let netBusinessRevenue = amountPaid - totalStaffCommission - businessProcessingFee - productTax;
    
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
      netBusinessRevenue = 0;
    }
    
    // Round function
    const roundToTwo = num => Math.round(num * 100) / 100;
    
    // Append a row to processedData
    processedData.push([
      trimmedTransactionID,                       // A: Transaction ID
      date,                                       // B: Date & Time
      serviceType,                                // C: Service Type
      staffName,                                  // D: Staff Member
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
      roundToTwo(netBusinessRevenue),             // U: Net Business Revenue
      statusRaw,                                  // V: Status
      customerName                                // W: Customer Name
    ]);
    
    Logger.log(`Processed Transaction ID: ${trimmedTransactionID}, Staff: ${staffName}, Customer: ${customerName}`);
  });
  
  // === Append processedData to "Processed Transactions" sheet ===
  if (processedData.length > 0) {
    const appendStartRow = processedLastRow + 1;
    processedSheet
      .getRange(appendStartRow, 1, processedData.length, headers.length)
      .setValues(processedData);
    Logger.log(`Appended ${processedData.length} new transactions to "Processed Transactions".`);
  } else {
    Logger.log(`No new transactions to process.`);
  }

  // === Sort "Processed Transactions" sheet by "Date & Time" (column B) descending ===
  const totalRows = processedSheet.getLastRow();
  if (totalRows > 1) {
    processedSheet
      .getRange(2, 1, totalRows - 1, headers.length)
      .sort({column: 2, ascending: false});
    Logger.log(`Sorted "Processed Transactions" sheet by "Date & Time" descending.`);
  }

  // === Reapply Formatting ===
  applyBusinessFormatting(processedSheet, headers.length);
  
  // === Remove Duplicate Transaction IDs ===
  removeDuplicateTransactionIDs();
}

/**
 * Removes duplicate rows in the "Processed Transactions" sheet based on "Transaction ID" (column A).
 */
function removeDuplicateTransactionIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const processedSheet = ss.getSheetByName('Processed Transactions');
  if (!processedSheet) throw new Error(`Sheet named 'Processed Transactions' not found.`);
  
  const lastRow = processedSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log(`No data to process in 'Processed Transactions' sheet.`);
    return;
  }
  
  const dataRange = processedSheet.getRange(2, 1, lastRow - 1, processedSheet.getLastColumn());
  const data = dataRange.getValues();
  const seenTransactionIDs = new Set();
  const rowsToDelete = [];
  
  // Iterate from bottom up so we delete duplicates after the first occurrence
  for (let i = data.length - 1; i >= 0; i--) {
    const transactionID = data[i][0]; // Column A
    if (transactionID) {
      const trimmedTransactionID = transactionID.toString().trim();
      if (seenTransactionIDs.has(trimmedTransactionID)) {
        rowsToDelete.push(i + 2); 
      } else {
        seenTransactionIDs.add(trimmedTransactionID);
      }
    }
  }
  
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => a - b);
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      processedSheet.deleteRow(rowsToDelete[i]);
    }
    Logger.log(`Removed ${rowsToDelete.length} duplicate row(s) based on "Transaction ID".`);
  } else {
    Logger.log(`No duplicate "Transaction ID" entries found.`);
  }
}

/**
 * Applies formatting to the "Processed Transactions" sheet (background colors, bold headers, auto-resize, etc.).
 * @param {Sheet} sheet 
 * @param {number} headerLength 
 */
function applyBusinessFormatting(sheet, headerLength) {
  const lastRow = sheet.getLastRow();
  
  // Background colors for processing fee columns G(7) & H(8): light blue
  const processingFeeColumns = [7, 8];
  processingFeeColumns.forEach(function(col){
    const range = sheet.getRange(1, col, lastRow);
    range.setBackground('#D9E1F2'); // Light blue
  });

  // Service-related columns I(9), J(10), K(11), L(12): light green
  const serviceColumns = [9, 10, 11, 12];
  serviceColumns.forEach(function(col){
    const range = sheet.getRange(1, col, lastRow);
    range.setBackground('#E2EFDA'); // Light green
  });

  // Product-related columns M(13), N(14), O(15), P(16), Q(17): light yellow
  const productColumns = [13, 14, 15, 16, 17];
  productColumns.forEach(function(col){
    const range = sheet.getRange(1, col, lastRow);
    range.setBackground('#FFF2CC'); // Light yellow
  });

  // Adjustment columns R(18) and S(19): light pink
  const adjustmentColumns = [18, 19];
  adjustmentColumns.forEach(function(col){
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

  Logger.log('Reapplied formatting to the "Processed Transactions" sheet.');
}

/**
 * Parses the service description field from raw data to extract staff name, service type, product info, etc.
 * @param {string} serviceDescription - The concatenated service description string from raw data
 * @param {number|string} quantityRaw - The raw quantity field
 * @returns {object} { staffName, serviceType, products[], additionalFees, nonProductCount }
 */
function parseServiceDescription(serviceDescription, quantityRaw) {
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
  const parts = serviceDescription.split(',').map(part => part.trim());
  
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