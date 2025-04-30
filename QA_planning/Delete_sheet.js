function updateStatusAndDeleteSheets() {
try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("delete"); // Change "delete" to the name of your sheet

  // Check if the source sheet exists
  if (!sheet) {
    Browser.msgBox("Sheet 'delete' not found");
    return;
  }

  // Get all the headers from the first row
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find column indexes dynamically
  var policyNameIndex = headers.indexOf("Policy Name");
  var statusIndex = headers.indexOf("Status");

  // If "Policy Name" header is not found, throw an error
  if (policyNameIndex === -1) {
    Browser.msgBox("Header 'Policy Name' not found");
    return;
  }

  // Add "Status" column if not found
  if (statusIndex === -1) {
    statusIndex = headers.length; // New column index (0-based)
    sheet.getRange(1, statusIndex + 1).setValue("Status");
  }

  // Increment indexes to 1-based for Google Sheets operations
  policyNameIndex += 1;
  statusIndex += 1;

  // Get data from the "Policy Name" column
  var data = sheet.getRange(2, policyNameIndex, sheet.getLastRow() - 1, 1).getValues();

  // Iterate over the data and update the "Status" column
  for (var i = 0; i < data.length; i++) {
    var sheetName = data[i][0];
    if (sheetName !== "") {
      var sheetToDelete = ss.getSheetByName(sheetName);
      if (sheetToDelete) {
        ss.deleteSheet(sheetToDelete);
        sheet.getRange(i + 2, statusIndex).setValue("Deleted");
      } else {
        sheet.getRange(i + 2, statusIndex).setValue("Not Found");
      }
    }
  }
 }
 catch (e) {
  handleError(e);
  }  
}

