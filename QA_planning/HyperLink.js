// Create hyperlink function with dynamic headers 
function create_sheet_hyper_link() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("QA_Report");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Sheet 'QA_Report' not found. Stopping the script.");
      return;
    }

    // Get header row and find column indices dynamically
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var slNoColIndex = headers.indexOf("Sl. No.") + 1;
    var policyNameColIndex = headers.indexOf("Policy Name") + 1;
    var sheetLinkColIndex = headers.indexOf("Sheet Link") + 1;

    // Check if all required headers are found
    if (slNoColIndex === 0 || policyNameColIndex === 0 || sheetLinkColIndex === 0) {
      SpreadsheetApp.getUi().alert("One or more required headers ('Sl. No.', 'Policy Name', 'Sheet Link') not found. Please check the headers.");
      return;
    }

    var lastRow = sheet.getLastRow();

    // Check if "Policy Name" column has any data
    var policyNameData = sheet.getRange(2, policyNameColIndex, lastRow - 1).getValues().flat();
    var hasData = policyNameData.some(name => name !== "");

    if (!hasData) {
      SpreadsheetApp.getUi().alert("No data found in the 'Policy Name' column. Stopping the script.");
      return;
    }

    // Process data in the "Policy Name" and "Sheet Link" columns
    for (var i = 2; i <= lastRow; i++) {
      // Set initial status in the "Sheet Link" column
      sheet.getRange(i, sheetLinkColIndex).setValue("...Fetching").setFontWeight("bold");

      // Get the Policy Name (sheet name reference)
      var rawSheetName = sheet.getRange(i, policyNameColIndex).getValue();

      if (rawSheetName) {
        var tmpSheet = ss.getSheetByName(rawSheetName);

        if (!tmpSheet) {
          sheet.getRange(i, sheetLinkColIndex).setValue("Sheet not found for " + rawSheetName).setFontWeight("normal");
        } else {
          var refSheetId = tmpSheet.getSheetId().toString();
          // Update hyperlink in "Sheet Link" column
          sheet.getRange(i, sheetLinkColIndex).setValue(['=HYPERLINK("#gid=' + refSheetId + '","' + rawSheetName + '")']).setFontWeight("normal");
        }
      }
    }

    // Add incremental formula for "Sl. No." column after hyperlinks are set
    for (var i = 2; i <= lastRow; i++) {
      sheet.getRange(i, slNoColIndex).setFormula(`=ROW() - 1`);
    }

    ss.toast("Number of records processed: " + (lastRow - 1), "👍 Process completed", 5);
  }
  catch (e) {
  handleError(e);
  logLibraryUsage('Create HyperLink', 'Fail', e.toString());
  Logger.log("Error: " + e.message); // Log the error for debugging purposes
 
  }
    
}

