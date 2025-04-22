function createSLAConfig_with_EmployeeCode_Optimized() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var separatedDataSheet = ss.getSheetByName("Seprated Data");
  var prodAssigneeConfigSheet = ss.getSheetByName("PROD Assignee Config");
  var slaConfigSheet = ss.getSheetByName("Ticketing_Form_vs._PROD_SLA");

  var separatedsheetheader = separatedDataSheet.getDataRange().getValues();
  var Prodsheetheader = prodAssigneeConfigSheet.getDataRange().getValues();

  var requiredHeaders = ['Category', 'Sub Categories'];
  var missingHeaders = [];

  for (var header of requiredHeaders) {
    if (separatedsheetheader[0].indexOf(header) === -1) {
      missingHeaders.push('Seprated File: ' + header);
    }
    if (Prodsheetheader[0].indexOf(header) === -1) {
      missingHeaders.push('PROD File: ' + header);
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
    slaConfigSheet = ss.insertSheet("Ticketing_Form_vs._PROD_SLA");
  }
  slaConfigSheet.setFrozenRows(1);

  var separatedData = separatedDataSheet.getDataRange().getValues();
  var prodAssigneeData = prodAssigneeConfigSheet.getDataRange().getValues();

  var separatedDataHeaders = separatedData[0];
  var prodAssigneeHeaders = prodAssigneeData[0];

  var prodAssigneeHeaderMap = {};
  for (var i = 0; i < prodAssigneeHeaders.length; i++) {
    prodAssigneeHeaderMap[prodAssigneeHeaders[i]] = i;
  }
ss.setActiveSheet(slaConfigSheet);
  slaConfigSheet.appendRow(["Category", "Sub Categories", "Form Name","Extra Field 1", "Assignee Email", "Group", "Escalation 1", "Escalation Time 1", "Escalation 2", "Escalation Time 2", "Seprate Row", "PROD assignee row", "Combination Status"]);

  var combinationsInSeparatedData = new Set();
  var combinationsInProdAssignee = new Set();

  var slaData = []; // For batch processing

  for (var i = 1; i < separatedData.length; i++) {
    var category = separatedData[i][separatedDataHeaders.indexOf("Category")];
    var subcategory = separatedData[i][separatedDataHeaders.indexOf("Sub Categories")];
    var FormName = separatedData[i][separatedDataHeaders.indexOf("Department")]; // in seprated data Department name contains Form name
    var employeeCode = separatedData[i][separatedDataHeaders.indexOf("Extra Field 1")];

    var combinationKey = category + subcategory + employeeCode;
    combinationsInSeparatedData.add(combinationKey);

    var matchFound = false;
    for (var j = 1; j < prodAssigneeData.length; j++) {
      if (
        category === prodAssigneeData[j][prodAssigneeHeaderMap["Category"]] &&
        subcategory === prodAssigneeData[j][prodAssigneeHeaderMap["Sub Categories"]] &&
        employeeCode === prodAssigneeData[j][prodAssigneeHeaderMap["Extra Field 1"]]
      ) {
        var rowData = [];
        rowData.push(category, subcategory, FormName, employeeCode);
        rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Assignee Email"]]);
        rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Group"]]);
        rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation 1"]]);
        rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation Time 1"]]);
        rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation 2"]]);
        rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation Time 2"]]);
        rowData.push(i + 1); // Line number in "Seprated Data"
        rowData.push(j + 1); // Line number in "PROD assignee config"
        rowData.push("Found in both files");
        slaData.push(rowData); // Batch processing
        matchFound = true;
        combinationsInProdAssignee.add(combinationKey);
        break;
      }
    }

    if (!matchFound) {
      var rowData = [];
      rowData.push(category, subcategory,FormName, employeeCode);
      rowData.push("Match not found");
      rowData.push("N/A");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push(i + 1); // Line number in "Seprated Data"
      rowData.push("N/A");
      rowData.push("Not found in PROD assignee config");
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

  for (var j = 1; j < prodAssigneeData.length; j++) {
    var category = prodAssigneeData[j][prodAssigneeHeaderMap["Category"]];
    var subcategory = prodAssigneeData[j][prodAssigneeHeaderMap["Sub Categories"]];
    var employeeCode = prodAssigneeData[j][prodAssigneeHeaderMap["Extra Field 1"]];

    var combinationKey = category + subcategory + employeeCode;
    if (!combinationsInProdAssignee.has(combinationKey)) {
      var rowData = [];
      rowData.push(category, subcategory, "N/A", employeeCode); // NA bcz form name is not available in PROD assignee
      rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Assignee Email"]]);
      rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Group"]]);
      rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation 1"]]);
      rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation Time 1"]]);
      rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation 2"]]);
      rowData.push(prodAssigneeData[j][prodAssigneeHeaderMap["Escalation Time 2"]]);
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
    var FormName = separatedData[i][separatedDataHeaders.indexOf("Department")]; // in seprated data Department name is Form name
    var employeeCode = separatedData[i][separatedDataHeaders.indexOf("Extra Field 1")];

    var combinationKey = category + subcategory + employeeCode;
    if (!combinationsInSeparatedData.has(combinationKey)) {
      var rowData = [];
      rowData.push(category, subcategory, FormName, employeeCode);
      rowData.push("Match not found");
      rowData.push("N/A");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push("Match not found");
      rowData.push(i + 1);
      rowData.push("N/A");
      rowData.push("Not found in PROD SLA config");
      extraRecordsInSeparatedData.push(rowData); // Batch processing
    }

    if (extraRecordsInSeparatedData.length >= 100) { // Batch size of 100 rows
      slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, extraRecordsInSeparatedData.length, extraRecordsInSeparatedData[0].length).setValues(extraRecordsInSeparatedData);
      extraRecordsInSeparatedData = []; // Reset for next batch
    }
  }

  if (extraRecordsInSeparatedData.length > 0) {
    slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, extraRecordsInSeparatedData.length, extraRecordsInSeparatedData[0].length).setValues(extraRecordsInSeparatedData);
    extraRecordsInSeparatedData = [];
  }
}
