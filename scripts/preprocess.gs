function processTransactionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName("Raw Data");
  var processedSheet = ss.getSheetByName("Processed Data");

  // Define the expected headers for the processed data
  var expectedHeaders = [
    "Date & Time", 
    "Transaction ID", 
    "Currency", 
    "Amount", 
    "Commission",
    "Processing Fee", 
    "Tip", 
    "Total Commission w/ Tip", 
    "Net Amount", 
    "Status",
    "Customer Name",
    "Order ID", 
    "Product Name", 
    "Quantity", 
    "Discount", 
    "Shipping", 
    "Tax"
  ];

  // === Step 1: Retrieve and Clean Raw Data ===
  var rawData = rawSheet.getDataRange().getValues();
  
  // Check if there are at least two rows to remove (assuming two header rows in raw data)
  if (rawData.length < 3) {
    Logger.log("Not enough data in raw data sheet.");
    return;
  }
  
  // Remove the first two rows (headers in raw data)
  rawData.splice(0, 2);

  // === Step 2: Retrieve Existing Transaction IDs from Processed Data ===
  var processedLastRow = processedSheet.getLastRow();
  var existingIDs = {};
  
  if (processedLastRow > 1) { 
    // Fetch Transaction IDs from Column B (2nd column) in processed data
    var processedIDs = processedSheet.getRange(2, 2, processedLastRow - 1, 1).getValues();
    processedIDs.forEach(function(row) {
      var id = row[0];
      if (id) {
        existingIDs[id] = true;
      }
    });
  }

  // === Step 3: Prepare New Data to Append ===
  var newProcessedData = [];

  rawData.forEach(function(row) {
    var transactionID = row[1]; // Column B in raw data (zero-based index 1)

    // Skip if Transaction ID already exists in processed data
    if (existingIDs[transactionID]) {
      return;
    }

    // Calculate commission and tip values
    var commission = parseFloat(row[6]) || 0; // Column G in raw data
    var tip = 0; // Can be modified as needed
    var totalCommissionWithTip = commission + tip;

    // Concatenate first name and last name from raw data
    var customerName = row[16] + " " + row[17]; // Columns Q and R

    var newRow = [
      row[0],               // [0]  Date & Time (A)
      row[1],               // [1]  Transaction ID (B)
      row[3],               // [2]  Currency (D)
      row[4],               // [3]  Amount (E)
      commission,           // [4]  Commission (calculated)
      row[5],               // [5]  Processing Fee (F)
      tip,                  // [6]  Tip
      totalCommissionWithTip, // [7] Total Commission w/ Tip
      row[7],               // [8]  Net Amount (H)
      row[8],               // [9]  Status (I)
      customerName,         // [10] Customer Name (concatenated)
      row[44],              // [11] Order ID (AS)
      row[45],              // [12] Product Name (AT)
      row[46],              // [13] Quantity (AU)
      row[47],              // [14] Discount (AV)
      row[48],              // [15] Shipping (AW)
      row[49]               // [16] Tax (AX)
    ];

    newProcessedData.push(newRow);
  });

  // If there's no new data to append, exit
  if (newProcessedData.length === 0) {
    Logger.log("No new data to append.");
    return;
  }

  // === Step 4: Ensure Processed Data Sheet has Correct Headers ===
  
  // Check if first row matches the expected headers
  function headersExist(sheet, headers) {
    var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (firstRow[i] !== headers[i]) {
        return false;
      }
    }
    return true;
  }

  // If the sheet is empty, add headers
  if (processedLastRow === 0) {
    processedSheet.appendRow(expectedHeaders);
    processedLastRow = 1;
  } else {
    // If there's at least one row, check if headers exist
    var headersPresent = headersExist(processedSheet, expectedHeaders);
    if (!headersPresent) {
      // Insert headers at the top
      processedSheet.insertRowBefore(1);
      processedSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      processedLastRow += 1;
    }
  }

  // === Step 5: Append New Data to Processed Sheet ===
  processedSheet
    .getRange(processedLastRow + 1, 1, newProcessedData.length, newProcessedData[0].length)
    .setValues(newProcessedData);
  
  // === Step 6: Sort the Processed Data by Date & Time in Descending Order ===
  var lastRow = processedSheet.getLastRow();
  var numRows = lastRow - 1; // Exclude header row
  if (numRows > 0) {
    processedSheet
      .getRange(2, 1, numRows, processedSheet.getLastColumn())
      .sort({column: 1, ascending: false});
  }

  Logger.log("Successfully appended " + newProcessedData.length + " new rows to processed data sheet.");
}