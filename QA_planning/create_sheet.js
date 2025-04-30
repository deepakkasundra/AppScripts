function create_sheets() {
  try {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("QA_Report");

 // Check if the sheet exists
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'QA_Report' not found. Stopping the script.");
    Logger.log("Sheet 'QA_Report' not found. Script execution stopped.");
    return; // Stop the script if the sheet is not found
  }


    // Get the header row to find the column index for "Policy Name"
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var policyNameColumnIndex = headerRow.indexOf("Policy Name"); // Dynamically find the index of "Policy Name"
    
    // Check if the column was found
    if (policyNameColumnIndex === -1) {
      throw new Error("The 'Policy Name' column was not found in the sheet.");
    }

    var lastRow = sheet.getLastRow();
  
    var dataRange = sheet.getRange(2, policyNameColumnIndex + 1, sheet.getLastRow() - 1); // Get the range for the "Policy Name" column starting from row 2
    var data = dataRange.getValues();

    // Check if there is any data in the Policy Name column
    if (data.length === 0 || (data.length === 1 && !data[0][0])) {
      var noDataMessage = "No data available in the Policy Name column. Please enter policy names before proceeding.";
      SpreadsheetApp.getActiveSpreadsheet().toast(noDataMessage, "Policy Name Check", 10);
      Browser.msgBox(noDataMessage, Browser.Buttons.OK); // Display a message box with the no data information
      return; // Exit the function if no data is available
    }    

    // Get values from the "Policy Name" column dynamically
    var sheetNames = sheet.getRange(2, policyNameColumnIndex + 1, lastRow - 1).getValues(); 

    // For each row in the sheet, insert a new sheet and rename it
    sheetNames.forEach(function(row) {
      if (row[0] != "") {
        var totalsheet = countSheets(); // Function to count existing sheets
        var sheetName = row[0];

        if (ss.getSheetByName(sheetName) != null) {
          ss.toast("Sheet already exists. Some records are skipped: " + sheetName, "⚠️ Warning", 5);
        } else {
          if (sheetName) { // Check if sheetName is not null or empty
            var indexSheet = ss.insertSheet(sheetName, totalsheet);
            ss.getSheetByName(sheetName).activate();
            ss.getRange('\'Test Cases Format\'!A1:P100').copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
          }
        }
      }
    });

    ss.toast("", "👍 Process completed", 5);
  } 
  catch (e) {
    handleError(e);
    Logger.log("Error: " + e.message); // Log the error for debugging purposes
  logLibraryUsage('Old Create Sheet', 'Fail', e.toString());
  }
  
}

