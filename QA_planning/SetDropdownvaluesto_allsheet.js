// Updated additional Dropdown options
function applyDataValidationToSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var options = ['Incorrect Response', 'Incomplete Response', 'incorrect policy in view source','PDF highlighting issue','Unclear response without view source']; // Replace with your dropdown options
  var validationRule = SpreadsheetApp.newDataValidation().requireValueInList(options).setAllowInvalid(false).build();
  var validationText = 'Select from ' + options.join(', ');

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
     Logger.log("Processing sheet: " + sheetName);
    if (sheetName !== "QA_Report" && sheetName !== "QA_Cycle_Wise_Report") {
      var range = sheet.getRange("F2:F");
      range.clearDataValidations(); // Clear existing data validations
      range.setDataValidation(validationRule);
      }
  }
}



// 'Incorrect Response', 'Incomplete Response', 'incorrect policy in view source','PDF highlighting issue','Unclear response without view source'
