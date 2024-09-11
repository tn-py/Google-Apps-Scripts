function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Raffle')
      .addItem('Calculate Entries Per Order', 'calculateEntries')
      .addItem('Summarize Entries Per Customer', 'summarizeEntries')
      .addItem('Generate All Entries', 'generateAllEntries')
      .addItem('Pick Raffle Winner', 'pickWinner')
      .addToUi();
  }
  
  
  function calculateEntries() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All Orders');
    var range = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1); // Getting the 'Amount' column
    var values = range.getValues();
  
    // Calculate entries per order and set in the 'Entries per Order' column (9th column)
    for (var i = 0; i < values.length; i++) {
      var amount = values[i][0];
      var entries = Math.floor(amount / 100); // 1 entry per $100
      sheet.getRange(i + 2, 9).setValue(entries); // Writing in the 'Entries per Order' column
    }
  
    SpreadsheetApp.getUi().alert('Entries per order calculated successfully.');
  }
  
  
  function summarizeEntries() {
    var ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All Orders');
    var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entries Per Customer');
    
    var ordersData = ordersSheet.getRange(2, 4, ordersSheet.getLastRow() - 1, 6).getValues(); // Get email and entries
    
    var customerEntries = {};
  
    // Summing up entries for each customer
    for (var i = 0; i < ordersData.length; i++) {
      var email = ordersData[i][0];
      var entries = ordersData[i][5]; // Entries Per Order
      if (!customerEntries[email]) {
        customerEntries[email] = 0;
      }
      customerEntries[email] += entries;
    }
  
    // Clearing previous data in the summary sheet
    summarySheet.getRange(2, 1, summarySheet.getLastRow(), 2).clearContent();
  
    // Writing the summarized entries per customer to the 'Entries Per Customer' sheet
    var row = 2;
    for (var email in customerEntries) {
      summarySheet.getRange(row, 1).setValue(email);
      summarySheet.getRange(row, 2).setValue(customerEntries[email]);
      row++;
    }
  
    SpreadsheetApp.getUi().alert('Entries per customer summarized successfully.');
  }
  
  
  function pickWinner() {
    var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entries Per Customer');
    var raffleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raffle');
  
    var entriesData = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 2).getValues(); // Get email and total entries
  
    var entriesList = [];
  
    // Create a list where each customer email appears according to their total entries
    for (var i = 0; i < entriesData.length; i++) {
      var email = entriesData[i][0];
      var totalEntries = entriesData[i][1];
  
      for (var j = 0; j < totalEntries; j++) {
        entriesList.push(email);
      }
    }
  
    if (entriesList.length === 0) {
      SpreadsheetApp.getUi().alert('No entries found. Cannot pick a winner.');
      return;
    }
  
    // Pick a random winner
    var randomIndex = Math.floor(Math.random() * entriesList.length);
    var winner = entriesList[randomIndex];
  
    // Write the winner to the 'Raffle' sheet
    raffleSheet.getRange(2, 1).setValue(winner);
  
    SpreadsheetApp.getUi().alert('Raffle winner selected: ' + winner);
  }
  
  
  function generateAllEntries() {
    var ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All Orders');
    var entriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All Entries');
    
    // Get all data from 'All Orders' (starting from row 2 to avoid headers)
    var ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 9).getValues(); 
    
    // Clear previous data in 'All Entries' sheet
    entriesSheet.getRange(2, 1, entriesSheet.getLastRow(), 4).clearContent();
    
    var allEntries = [];
    
    // Loop through 'All Orders' to extract entries
    for (var i = 0; i < ordersData.length; i++) {
      var documentNumber = ordersData[i][1]; // Document Number
      var name = ordersData[i][2];           // Customer Name
      var email = ordersData[i][3];          // Email
      var entriesPerOrder = ordersData[i][8]; // Entries per Order
      
      // Generate one row per entry with "1" in the Entry column
      for (var j = 0; j < entriesPerOrder; j++) {
        allEntries.push([documentNumber, name, email, 1]); // Create an entry row with "1" in the 'Entry' column
      }
    }
  
    // Insert all entries into the 'All Entries' sheet starting from row 2
    if (allEntries.length > 0) {
      entriesSheet.getRange(2, 1, allEntries.length, 4).setValues(allEntries);
    }
  
    SpreadsheetApp.getUi().alert('All Entries have been generated successfully.');
  }
  
  
  