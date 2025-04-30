
function set_workLM() {
  try {
  set_sheet("worklm");
  }
 catch (e) {
  handleError(e);
  }  
}

function set_Autonomous() {
try{
  set_sheet("autonomous");
}
 catch (e) {
  handleError(e);
  }  
}

function set_sheet(option) {
    try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qaReportSheet = ss.getSheetByName("QA_Report");
  var headers = qaReportSheet.getDataRange().getValues()[0]; // Assuming headers are in the first row
  unhideAllColumns(qaReportSheet);

  if (!qaReportSheet) {
    SpreadsheetApp.getUi().alert("QA_Report sheet not found.");
    Logger.log("QA_Report sheet not found.");
    return;
  }

  if (option.toLowerCase() === "worklm") {
  
    // Unhide "WorkLM_Test Cases Format" and hide "Autonomous_Test Cases Format"
    var worklmSheet = ss.getSheetByName("WorkLM_Test Cases Format");
    var autonomousSheet = ss.getSheetByName("Autonomous_Test Cases Format");

    if (worklmSheet) worklmSheet.showSheet();
    if (autonomousSheet) autonomousSheet.hideSheet();

    // Hide specific columns in "QA_Report"
    hideColumns(qaReportSheet, headers, [
      "Bot stuck",
      "Incorrect audience",
      "Incorrect response",
      "hallucination issue",
      "Policy search",
      "No/Fallback response",
      "Partial response given",
      "Generic response",
      "Incorrect language",
      "Incorrect Highlighting"
    ]);

   ss.toast("Project Sheet set as WorkLM/GPT ", "👍 Project set successfully", 5);
  logLibraryUsage('Set Work Sheet as Worklm', 'Pass');  
  
  } else if (option.toLowerCase() === "autonomous") {

    // Unhide "Autonomous_Test Cases Format" and hide "WorkLM_Test Cases Format"
    var worklmSheet = ss.getSheetByName("WorkLM_Test Cases Format");
    var autonomousSheet = ss.getSheetByName("Autonomous_Test Cases Format");

    if (autonomousSheet) autonomousSheet.showSheet();
    if (worklmSheet) worklmSheet.hideSheet();

    // Hide specific columns in "QA_Report"
    hideColumns(qaReportSheet, headers, [
      "Response Count from GPT",
      "% Response from GPT",
      "GPT Pass %",
      "GPT Fail %",
      "Fallback Count",
      "Fallback %",
      "response from search count",
      "% response from Search",
      "GPT Pass Count",
      "GPT Fail Count"
    ]);

    ss.toast("Project Sheet set as Autonomous ", "👍 Project set successfully", 5);
  logLibraryUsage('Set Work Sheet as Autonomous', 'Pass');  
 
  } else {
    SpreadsheetApp.getUi().alert("Invalid option provided. Please use 'worklm' or 'autonomous'.");
    Logger.log("Invalid option: " + option);
    return;
  }
}

 catch (e) {
  handleError(e);
  logLibraryUsage('Set Work Sheet', 'Fail', e.toString());
  }
}
