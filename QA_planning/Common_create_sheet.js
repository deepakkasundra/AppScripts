function create_sheet_workLM() {
  create_sheets_common("worklm");
}

function create_sheet_Autonomous() {
  create_sheets_common("autonomous");
}

function create_sheets_common(option) {
  try {

            // Set the number of rows to insert in new sheet
          var CONFIG_ROW_COUNT = 100;

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
     // Set the formula in the "Sr. No" column
    // if (lastRow > 1) { // Ensure there is data to apply the formula
    //   var srNoRange = sheet.getRange(2, srNoColumnIndex + 1, lastRow - 1);
    //   srNoRange.setFormula(`=ARRAYFORMULA(IF(ROW(A2:A) - ROW(A$2) + 1 <= COUNTA(B2:B), ROW(A2:A) - ROW(A$2) + 1, ""))`);
    // }
    
    var dataRange = sheet.getRange(2, policyNameColumnIndex + 1, sheet.getLastRow() - 1); // Get the range for the "Policy Name" column starting from row 2
    var data = dataRange.getValues();

    // Check if there is any data in the Policy Name column
    if (data.length === 0 || (data.length === 1 && !data[0][0])) {
      var noDataMessage = "No data available in the Policy Name column. Please enter policy names before proceeding.";
      SpreadsheetApp.getActiveSpreadsheet().toast(noDataMessage, "Policy Name Check", 10);
      Browser.msgBox(noDataMessage, Browser.Buttons.OK); // Display a message box with the no data information
      return; // Exit the function if no data is available
    }    

    // Determine the template sheet based on the selected option
    var templateSheetName;
    if (option.toLowerCase() === "worklm") {
      templateSheetName = "WorkLM_Test Cases Format";
    } else if (option.toLowerCase() === "autonomous") {
      templateSheetName = "Autonomous_Test Cases Format";
    } else {
      SpreadsheetApp.getUi().alert("Invalid option provided. Please use 'worklm' or 'autonomous'.");
      Logger.log("Invalid option: " + option);
      return;
    }
    
// Check if the template sheet is hidden
var templateSheet = ss.getSheetByName(templateSheetName);
if (templateSheet.isSheetHidden()) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "The selected Project Type doesn't match the template sheet '" + templateSheetName + "'. Please check and select the correct Options.",    
    ui.ButtonSet.YES_NO
  );

  // If the user clicked "No" or dismissed the dialog
  if (response !== ui.Button.YES) { 
    Logger.log("User chose to stop due to hidden template sheet or dismissed the alert.");
    ss.toast("User chose to stop due to hidden template sheet."," ⚠️ Process Terminated",10);
    return; // Stop the script if user chooses "No" or dismisses
  }

  // Unhide the template sheet if the user explicitly selects "Yes"
  templateSheet.showSheet();
  Logger.log("Template sheet '" + templateSheetName + "' has been unhidden.");
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

            indexSheet.deleteRows(CONFIG_ROW_COUNT + 1, indexSheet.getMaxRows() - CONFIG_ROW_COUNT);
            ss.getSheetByName(sheetName).activate();
            ss.getRange(`'${templateSheetName}'!A1:P100`).copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);
          }
        }
      }
    });

var hidetemplatesheet;

    if (option.toLowerCase() === "worklm") {
      hidetemplatesheet = "Autonomous_Test Cases Format";
    } else if (option.toLowerCase() === "autonomous") {
      hidetemplatesheet = "WorkLM_Test Cases Format";
    } else {
      SpreadsheetApp.getUi().alert("Invalid option provided. Please use 'worklm' or 'autonomous'.");
      Logger.log("Invalid option: " + option);
      return;
    }

 ss.getSheetByName(hidetemplatesheet).hideSheet();
    ss.toast("", "👍 Process completed", 5);
  }
   catch (e) {
  handleError(e);
  logLibraryUsage('Set Work Sheet', 'Fail', e.toString()); 
  }  
  }


