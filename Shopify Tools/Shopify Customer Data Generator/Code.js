function generateSampleData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var numRows = promptForRows();
  
  var data = [];
  for (var i = 0; i < numRows; i++) {
    data.push(generateRow());
  }
  
  sheet.getRange(2, 1, numRows, data[0].length).setValues(data);
}

function promptForRows() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('How many rows of sample customers would you like?', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var numRows = parseInt(response.getResponseText());
    if (!isNaN(numRows) && numRows > 0) {
      return numRows;
    } else {
      ui.alert('Please enter a valid number.');
      return promptForRows();
    }
  } else {
    return 0; // Cancel was clicked or some other non-OK response
  }
}

function generateRow() {
  return [
    // First Name
    generateFirstName(),
    // Last Name
    generateLastName(),
    // Email
    generateEmail(),
    // Accepts Email Marketing
    generateBoolean(),
    // Default Address Company
    "Sample Company",
    // Default Address Address1
    "123 Sample St",
    // Default Address Address2
    "Suite 100",
    // Default Address City
    "Sample City",
    // Default Address Province Code
    "SC",
    // Default Address Country Code
    "US",
    // Default Address Zip
    "12345",
    // Default Address Phone
    generatePhone(),
    // Phone
    generatePhone(),
    // Accepts SMS Marketing
    generateBoolean(),
    // Tags
    "Sample Tag",
    // Note
    "Sample Note",
    // Tax Exempt
    generateBoolean()
  ];
}

function generateFirstName() {
  var firstNames = ["John", "Jane", "Mike", "Sally", "Laura"];
  return firstNames[Math.floor(Math.random() * firstNames.length)];
}

function generateLastName() {
  var lastNames = ["Doe", "Smith", "Johnson", "Williams", "Brown"];
  return lastNames[Math.floor(Math.random() * lastNames.length)];
}

function generateEmail() {
  return generateFirstName().toLowerCase() + "." + generateLastName().toLowerCase() + "@example.com";
}

function generateBoolean() {
  return Math.random() < 0.5 ? "Yes" : "No";
}

function generatePhone() {
  var phone = "+1";
  for (var i = 0; i < 10; i++) {
    phone += Math.floor(Math.random() * 10);
  }
  return phone;
}
