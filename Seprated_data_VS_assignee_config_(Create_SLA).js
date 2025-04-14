// Seprated_data_VS_assignee_config_(Create_SLA)
function fetchWithEmployeeCode_UAT() {
  try {
    createSLAConfig_Optimized(true, 'UAT Assignee Config', 'Ticketing_Form_vs_UAT_SLA', "Seprated vs. UAT With Employee Code");
  } catch (error) {
    Logger.log("Error in fetchWithEmployeeCode_UAT: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
  }
}

function fetchWithoutEmployeeCode_UAT() {
  try {
    createSLAConfig_Optimized(false, 'UAT Assignee Config', 'Ticketing_Form_vs_UAT_SLA', "Seprated vs. UAT Without Employee Code");
  } catch (error) {
    Logger.log("Error in fetchWithoutEmployeeCode_UAT: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
  }
}

function fetchWithEmployeeCode_PROD() {
  try {
    createSLAConfig_Optimized(true, 'PROD Assignee Config', 'Ticketing_Form_vs_PROD_SLA', "Seprated vs. PROD With Employee Code");
  } catch (error) {
    Logger.log("Error in fetchWithEmployeeCode_PROD: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
  }
}

function fetchWithoutEmployeeCode_PROD() {
  try {
    createSLAConfig_Optimized(false, 'PROD Assignee Config', 'Ticketing_Form_vs_PROD_SLA', "Seprated vs. PROD Without Employee Code");
  } catch (error) {
    Logger.log("Error in fetchWithoutEmployeeCode_PROD: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
  }
}

function fetchWithEmployeeCode_REQ() {
  try {
    createSLAConfig_Optimized(true, 'Requirement Assignee Config', 'Ticketing_Form_vs_Requirement_SLA', "Seprated vs. Requirement With Employee Code");
  } catch (error) {
    Logger.log("Error in fetchWithEmployeeCode_REQ: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
  }
}

function fetchWithoutEmployeeCode_REQ() {
  try {
    createSLAConfig_Optimized(false, 'Requirement Assignee Config', 'Ticketing_Form_vs_Requirement_SLA', "Seprated vs. Requirement Without Employee Code");
  } catch (error) {
    Logger.log("Error in fetchWithoutEmployeeCode_REQ: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
  }
}

function createSLAConfig_Optimized(useEmployeeCode, assigneeConfigSheetName, resultSheetName, selectedOption) {
try {

  showProgressToast(ss, "Started processing for: " + selectedOption);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var separatedDataSheet = ss.getSheetByName("Seprated Data");
  var assigneeConfigSheet = ss.getSheetByName(assigneeConfigSheetName);
  var slaConfigSheet = ss.getSheetByName(resultSheetName);

  var separatedsheetheader = separatedDataSheet.getDataRange().getValues();
  var assigneeSheetHeader = assigneeConfigSheet.getDataRange().getValues();

  var requiredHeaders = ['Category', 'Sub Categories'];
  var missingHeaders = [];

  for (var header of requiredHeaders) {
    if (separatedsheetheader[0].indexOf(header) === -1) {
      missingHeaders.push('Seprated File: ' + header);
    }
    if (assigneeSheetHeader[0].indexOf(header) === -1) {
      missingHeaders.push(assigneeConfigSheetName + ': ' + header);
    }
  }

  if (missingHeaders.length > 0) {
    var missingHeadersString = missingHeaders.join(', ');
    SpreadsheetApp.getActiveSpreadsheet().toast('The following required column headers are missing: ' + missingHeadersString, '⚠️ Warning', 20);
    return;
  }

  if (slaConfigSheet) {
    slaConfigSheet.clear();
  } else {
    slaConfigSheet = ss.insertSheet(resultSheetName);
  }
  slaConfigSheet.setFrozenRows(1);
ss.setActiveSheet(slaConfigSheet);
  var separatedData = separatedDataSheet.getDataRange().getValues();
  var assigneeData = assigneeConfigSheet.getDataRange().getValues();



  var separatedDataHeaders = separatedData[0];
  var assigneeHeaders = assigneeData[0];

  var assigneeHeaderMap = {};
  for (var i = 0; i < assigneeHeaders.length; i++) {
    assigneeHeaderMap[assigneeHeaders[i]] = i;
  }

  slaConfigSheet.appendRow(["Category", "Sub Categories", "Department Name", "Form Name","Extra Field 1", "Assignee Email", "Group", "Escalation 1", "Escalation Time 1", "Escalation 2", "Escalation Time 2", "Seprate Row", "Assignee Row", "Combination Status"]);

  var combinationsInSeparatedData = new Set();
  var combinationsInAssignee = new Set();

  var slaData = []; // For batch processing

  for (var i = 1; i < separatedData.length; i++) {
    var category = separatedData[i][separatedDataHeaders.indexOf("Category")];
    var subcategory = separatedData[i][separatedDataHeaders.indexOf("Sub Categories")];
    var Department = separatedData[i][separatedDataHeaders.indexOf("Department")]; 
    // in separated data Department name contains Form name
    var employeeCode = separatedData[i][separatedDataHeaders.indexOf("Extra Field 1")];
    var FormName = separatedData[i][separatedDataHeaders.indexOf("Form Name")];

  // Combination added as Department + Cat + Sub
    var combinationKey = useEmployeeCode ? (Department + category + subcategory + employeeCode) : (Department + category + subcategory);
    combinationsInSeparatedData.add(combinationKey);

    var matchFound = false;
    for (var j = 1; j < assigneeData.length; j++) {
      var assigneeDepartment = assigneeData[j][assigneeHeaderMap["Department"]];
      var assigneeCategory = assigneeData[j][assigneeHeaderMap["Category"]];
      var assigneeSubcategory = assigneeData[j][assigneeHeaderMap["Sub Categories"]];
      var assigneeEmployeeCode = assigneeData[j][assigneeHeaderMap["Extra Field 1"]];

      if (
        Department === assigneeDepartment &&
        category === assigneeCategory &&
        subcategory === assigneeSubcategory &&
        (useEmployeeCode ? employeeCode === assigneeEmployeeCode : true)
      ) {
        var rowData = [];
        rowData.push(category, subcategory, Department, FormName,employeeCode);
        rowData.push(assigneeData[j][assigneeHeaderMap["Assignee Email"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Group"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 1"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 1"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 2"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 2"]]);
        rowData.push(i + 1); // Line number in "Separated Data"
        rowData.push(j + 1); // Line number in assignee config
        rowData.push("Found in both files");
        slaData.push(rowData); // Batch processing
        matchFound = true;
        combinationsInAssignee.add(combinationKey);
        break;
      }
    }

    if (!matchFound) {
      var rowData = [];
      rowData.push(category, subcategory, Department, FormName,employeeCode);
      rowData.push("Match not found");
      rowData.push("N/A");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push(i + 1); // Line number in "Separated Data"
      rowData.push("N/A");
      rowData.push("Not found in assignee config");
      slaData.push(rowData); // Batch processing
    }

    if (slaData.length >= 100) { // Batch size of 100 rows
      slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, slaData.length, slaData[0].length).setValues(slaData);
      slaData = []; // Reset for next batch
    }
  }

  if (slaData.length > 0) {
    slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, slaData.length, slaData[0].length).setValues(slaData);
    slaData = [];
  }

  for (var j = 1; j < assigneeData.length; j++) {
    var assigneeDepartment = assigneeData[j][assigneeHeaderMap["Department"]];
    var assigneeCategory = assigneeData[j][assigneeHeaderMap["Category"]];
    var assigneeSubcategory = assigneeData[j][assigneeHeaderMap["Sub Categories"]];
    var assigneeEmployeeCode =  assigneeData[j][assigneeHeaderMap["Extra Field 1"]];
    // combination updated as Department + Cate + sub
    var combinationKey = assigneeDepartment + assigneeCategory + assigneeSubcategory + (useEmployeeCode ? assigneeEmployeeCode : '');
    
      if (!combinationsInAssignee.has(combinationKey)) {
      var rowData = [];
      rowData.push(assigneeCategory, assigneeSubcategory, assigneeDepartment, "N/A",assigneeEmployeeCode ); 
      // NA because form name is not available in assignee config
      rowData.push(assigneeData[j][assigneeHeaderMap["Assignee Email"]]);
      rowData.push(assigneeData[j][assigneeHeaderMap["Group"]]);
      rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 1"]]);
      rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 1"]]);
      rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 2"]]);
      rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 2"]]);
      rowData.push("N/A");
      rowData.push(j + 1);
      rowData.push("Not found in Ticketing Form");
      slaData.push(rowData); // Batch processing
    }

    if (slaData.length >= 100) { // Batch size of 100 rows
      slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, slaData.length, slaData[0].length).setValues(slaData);
      slaData = []; // Reset for next batch
    }
  }

  if (slaData.length > 0) {
    slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, slaData.length, slaData[0].length).setValues(slaData);
    slaData = [];
  }

  var extraRecordsInSeparatedData = [];
  for (var i = 1; i < separatedData.length; i++) {
    var category = separatedData[i][separatedDataHeaders.indexOf("Category")];
    var subcategory = separatedData[i][separatedDataHeaders.indexOf("Sub Categories")];
    var Department = separatedData[i][separatedDataHeaders.indexOf("Department")]; // in separated data department name is form name
    var employeeCode = separatedData[i][separatedDataHeaders.indexOf("Extra Field 1")];

    var combinationKey = useEmployeeCode ? (Department + category + subcategory + employeeCode) : (Department + category + subcategory);
    
    if (!combinationsInSeparatedData.has(combinationKey)) {
      var rowData = [];
      rowData.push(category, subcategory, Department, FormName, employeeCode );
      rowData.push("Match not found");
      rowData.push("N/A");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push(i + 1);
      rowData.push("N/A");
      rowData.push("Not found in assignee config");
      extraRecordsInSeparatedData.push(rowData);
    }

    if (extraRecordsInSeparatedData.length >= 100) { // Batch size of 100 rows
      slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, extraRecordsInSeparatedData.length, extraRecordsInSeparatedData[0].length).setValues(extraRecordsInSeparatedData);
      extraRecordsInSeparatedData = []; // Reset for next batch
    }
  }

  if (extraRecordsInSeparatedData.length > 0) {
    slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, extraRecordsInSeparatedData.length, extraRecordsInSeparatedData[0].length).setValues(extraRecordsInSeparatedData);
  }

  showProgressToast(ss, "Completed: " + selectedOption);
} catch (error) {
    Logger.log("Error in createSLAConfig_Optimized: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in processing: " + error.message, '⚠️ Error', 5);
  }
}

