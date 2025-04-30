function updateQACycleReport() {
  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qaReportSheet = ss.getSheetByName("QA_Report");
  var qaCycleReportSheet = ss.getSheetByName("QA_Cycle_Wise_Report");

  // Check if both sheets exist
  if (!qaReportSheet || !qaCycleReportSheet) {
    SpreadsheetApp.getUi().alert("One or both of the sheets ('QA_Report' or 'QA_Cycle_Wise_Report') are missing. Please check the sheet names.");
    return;
  }

  // Define the columns to sum
  var columnsToSum = ["Test Cases Count", "Pass", "Fail", "Response Count from GPT", "GPT Pass Count"];

  // Get the headers from both sheets
  var qaReportHeaders = qaReportSheet.getRange(1, 1, 1, qaReportSheet.getLastColumn()).getValues()[0];
  var qaCycleReportHeaders = qaCycleReportSheet.getRange(1, 1, 1, qaCycleReportSheet.getLastColumn()).getValues()[0];

// Check if "Date Time" column is present, if not add it
    if (qaCycleReportHeaders.indexOf("Date Time") === -1) {
      qaCycleReportSheet.insertColumnAfter(qaCycleReportHeaders.length);
      qaCycleReportSheet.getRange(1, qaCycleReportHeaders.length + 1).setValue("Date Time");
      qaCycleReportHeaders.push("Date Time");
    }
    
  // Get the data from QA_Report sheet
  var qaReportData = qaReportSheet.getDataRange().getValues();

  // Get the last cycle number from QA_Cycle_Wise_Report
  var qaCycleReportData = qaCycleReportSheet.getDataRange().getValues();
  var lastQACycle = qaCycleReportData.length > 1 ? parseInt(qaCycleReportData[qaCycleReportData.length - 1][0]) : 0;

  // Initialize the new QA cycle value
  var newQACycle = qaCycleReportData.length === 0 ? 1 : lastQACycle + 1;

  // Initialize an object to store the sums
  var sums = { "QA Cycle": newQACycle };

  // Get the current date and time formatted as DD/MM/YYYY HH:MM:SS
  var currentDateTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");


  // Calculate the sum for each respective column dynamically
  columnsToSum.forEach(columnName => {
    var columnIndex = qaReportHeaders.indexOf(columnName);

    if (columnIndex !== -1) {
      var columnValues = qaReportSheet.getRange(2, columnIndex + 1, qaReportSheet.getLastRow() - 1, 1).getValues();
      var sum = columnValues.reduce((a, b) => a + (b[0] || 0), 0); // Handle empty cells
      sums[columnName] = sum;
    } else {
      sums[columnName] = ""; // If the column is not found, add an empty value
    }
  });

  // Prepare the row to append
  var rowToAppend = [];
  qaCycleReportHeaders.forEach(header => {
    if (header === "Date Time") {
      rowToAppend.push(currentDateTime); // Add the current date and time
    } else {
      rowToAppend.push(sums[header] !== undefined ? sums[header] : "");
    }
  });

  // Append the row to QA_Cycle_Wise_Report
  qaCycleReportSheet.appendRow(rowToAppend);

  // Search for the header "GPT Pass %" and set the formula in the respective column
  var gptPassPercentIndex = qaCycleReportHeaders.indexOf("GPT Pass %");

  if (gptPassPercentIndex !== -1) {
    var gptPassCountIndex = qaCycleReportHeaders.indexOf("GPT Pass Count");
    var responseCountIndex = qaCycleReportHeaders.indexOf("Response Count from GPT");

    if (gptPassCountIndex !== -1 && responseCountIndex !== -1) {
      var formulaCell = qaCycleReportSheet.getRange(lastQACycle + 2, gptPassPercentIndex + 1);
      var formula = `=IFERROR(${String.fromCharCode(65 + gptPassCountIndex)}${lastQACycle + 2}/${String.fromCharCode(65 + responseCountIndex)}${lastQACycle + 2}*100, "")`;
      formulaCell.setFormula(formula);
    } else {
      SpreadsheetApp.getUi().alert("Required columns for GPT Pass % calculation are missing.");
    }
  }

  // Notify the user that the process is complete
  ss.toast("QA_Cycle_Wise_Report sheet Updated", "👍 Process completed", 5);
} 
catch (e) {
  handleError(e);
  logLibraryUsage('Old QA Cycle Report', 'Fail', e.toString());
  Logger.log("Error: " + e.message); // Log the error for debugging purposes
 
  }
  
}


// function updateQACycleReport() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var qaReportSheet = ss.getSheetByName("QA_Report");
//   var qaCycleReportSheet = ss.getSheetByName("QA_Cycle_Wise_Report");

//   // Define the columns to sum ("Test Case Count", Pass, Fail, "Response Count from GPT", "GPT Pass Count", and "GPT Fail Count")
//   var columnsToSum = ["Test Cases Count", "Pass", "Fail", "Response Count from GPT", "GPT Pass Count"];

//   // Get the data from QA_Report sheet
//   var qaReportData = qaReportSheet.getDataRange().getValues();

//   // Get the data from QA_Cycle_Wise_Report sheet
//   var qaCycleReportData = qaCycleReportSheet.getDataRange().getValues();
//   var lastQACycle = qaCycleReportData.length > 1 ? parseInt(qaCycleReportData[qaCycleReportData.length - 1][0]) : 0;

//   // Initialize the QA cycle value to 1 if the QA_Cycle_Wise_Report sheet is empty
//   var newQACycle = qaCycleReportData.length === 0 ? 1 : lastQACycle + 1;

//   // Initialize an array to store the sums
//   var sums = [newQACycle];

//   // Calculate sum for each respective column (including "Response Count from GPT", "GPT Pass Count", and "GPT Fail Count")
//   for (var i = 0; i < columnsToSum.length; i++) {
//     var column = columnsToSum[i];
//     var columnIndex = qaReportSheet.getRange(1, 1, 1, qaReportSheet.getLastColumn()).getValues()[0].indexOf(column);

//     if (columnIndex !== -1) {
//       var columnValues = qaReportSheet.getRange(2, columnIndex + 1, qaReportSheet.getLastRow() - 1, 1).getValues();
//       var sum = columnValues.reduce(function(a, b) { return a + b[0]; }, 0); // Sum the values in the column
//       sums.push(sum);
//     } else {
//       sums.push(""); // If the column is not found, add an empty cell
//     }
    
//   }

//   // Append the sums to QA_Cycle_Wise_Report
//   qaCycleReportSheet.appendRow(sums);


// // Search for the header "GPT Pass %" and set the formula in the cell below the header
//   var headerRow = qaCycleReportData[0];
  
//   for (var j = 0; j < headerRow.length; j++) {
//     if (headerRow[j] === "GPT Pass %") {
//       var columnLetter = String.fromCharCode(65 + j); // Convert column index to letter (A=0, B=1, ...)
//       var formulaCell = qaCycleReportSheet.getRange(lastQACycle + 2, j + 1); // Move one row below the header
//       var formula = `=IFERROR(F${lastQACycle + 2}/E${lastQACycle + 2}*100, "")`; // Assuming "Response Count from GPT" is in column D
//       formulaCell.setFormula(formula);
//       break;
//     }
//   }
//   }

