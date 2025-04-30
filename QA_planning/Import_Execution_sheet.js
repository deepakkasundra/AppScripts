function mapDataToSheets() {
 try {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = "Execution"; // Change to your source sheet name
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);

  // Check if the source sheet exists
  if (!sourceSheet) {
    Browser.msgBox("Source sheet '" + sourceSheetName + "' not found.");
    return;
  }

  // Get the data range in the source sheet
  var dataRange = sourceSheet.getDataRange();
  var values = dataRange.getValues();

  // Initialize an array to track cleared sheets
  var clearedSheets = [];

  // Loop through the data and create/modify sheets
  for (var i = 1; i < values.length; i++) { // Assuming data starts from row 2
    var policyName = values[i][0]; // Assuming Policy Name is in the first column (column 0)

    var destinationSheet = spreadsheet.getSheetByName(policyName);

    // If the destination sheet doesn't exist, mark the "Process_record_status" column as "Sheet not found" for this record
    if (destinationSheet == null) {
      sourceSheet.getRange(i + 1, 11).setValue("Sheet not found");
      continue; // Skip processing if the sheet doesn't exist
    }

   
    // Clear existing data in the destination sheet only if it hasn't been cleared before
    if (!clearedSheets.includes(policyName)) {
      clearDestinationSheet(destinationSheet);
      clearedSheets.push(policyName); // Add the sheet name to the cleared sheets array
    }

    // Get the last used "Sr no" in the destination sheet
    var lastSrNo = destinationSheet.getLastRow() - 1; // Subtract 1 to account for the header row
    var nextSrNo = isNaN(lastSrNo) ? 1 : lastSrNo + 1;

    // Map data to the destination sheet, starting "Sr no" from nextSrNo
    var rowData = [
      nextSrNo, // Sr no starts from 1 in each sheet
      values[i][1], // Testcase
      values[i][2], // Page No
      values[i][3], // Process_record_status (update this column in source sheet)
      values[i][4], // Direct
      "", // Column 6 - Leave it empty
      values[i][9] // GPT Response
    ];

    destinationSheet.appendRow(rowData);

    // Update the "Process_record_status" in the source sheet as "data pass" for this record
    sourceSheet.getRange(i + 1, 11).setValue("Data Cpoied");

    // Display a progress message with the policy name
    var progressMessage = "Processing record " + i + " out of " + (values.length - 1) + " for Policy: " + policyName;
    spreadsheet.toast(progressMessage, "Progress", 10);
  }

  // Display a completion message when all records are processed
  spreadsheet.toast("Processing completed!" , "Progress", -1);
  }
 catch (e) {
  handleError(e);
  }  

}

// Function to clear existing data in the destination sheet except for the header row
function clearDestinationSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}


