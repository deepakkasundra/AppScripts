	//sub category as optional Seprated_data_VS_assignee_config_(Create_SLA)
	function fetchWithEmployeeCode_UAT_subcat_optional() {
	  try {
		createSLAConfig_Optimized_subcat_optional(true, 'UAT Assignee Config', 'Ticketing_Form_vs_UAT_SLA', "Seprated vs. UAT With Employee Code");
	  } catch (error) {
		Logger.log("Error in fetchWithEmployeeCode_UAT_subcat_optional: " + error.message);
		SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
	  }
	}

	function fetchWithoutEmployeeCode_UAT_subcat_optional() {
	  try {
		createSLAConfig_Optimized_subcat_optional(false, 'UAT Assignee Config', 'Ticketing_Form_vs_UAT_SLA', "Seprated vs. UAT Without Employee Code");
	  } catch (error) {
		Logger.log("Error in fetchWithoutEmployeeCode_UAT_subcat_optional: " + error.message);
		SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
	  }
	}

	function fetchWithEmployeeCode_PROD_subcat_optional() {
	  try {
		createSLAConfig_Optimized_subcat_optional(true, 'PROD Assignee Config', 'Ticketing_Form_vs_PROD_SLA', "Seprated vs. PROD With Employee Code");
	  } catch (error) {
		Logger.log("Error in fetchWithEmployeeCode_PROD_subcat_optional: " + error.message);
		SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
	  }
	}

	function fetchWithoutEmployeeCode_PROD_subcat_optional() {
	  try {
		createSLAConfig_Optimized_subcat_optional(false, 'PROD Assignee Config', 'Ticketing_Form_vs_PROD_SLA', "Seprated vs. PROD Without Employee Code");
	  } catch (error) {
		Logger.log("Error in fetchWithoutEmployeeCode_PROD_subcat_optional: " + error.message);
		SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
	  }
	}

	function fetchWithEmployeeCode_REQ_subcat_optional() {
	  try {
		createSLAConfig_Optimized_subcat_optional(true, 'Requirement Assignee Config', 'Ticketing_Form_vs_Requirement_SLA', "Seprated vs. Requirement With Employee Code");
	  } catch (error) {
		Logger.log("Error in fetchWithEmployeeCode_REQ_subcat_optional: " + error.message);
		SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
	  }
	}

	function fetchWithoutEmployeeCode_REQ_subcat_optional() {
	  try {
		createSLAConfig_Optimized_subcat_optional(false, 'Requirement Assignee Config', 'Ticketing_Form_vs_Requirement_SLA', "Seprated vs. Requirement Without Employee Code");
	  } catch (error) {
		Logger.log("Error in fetchWithoutEmployeeCode_REQ_subcat_optional: " + error.message);
		SpreadsheetApp.getActiveSpreadsheet().toast("Error in fetching data: " + error.message, '⚠️ Error', 5);
	  }
	}

	// Function to create SLA configuration with optimized logic
function createSLAConfig_Optimized_subcat_optional(useEmployeeCode, assigneeConfigSheetName, resultSheetName, selectedOption) {
  try {
    showProgressToast(SpreadsheetApp.getActiveSpreadsheet(), "Started processing for: " + selectedOption);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var separatedDataSheet = ss.getSheetByName("Seprated Data");
    var assigneeConfigSheet = ss.getSheetByName(assigneeConfigSheetName);
    var slaConfigSheet = ss.getSheetByName(resultSheetName);

    var separatedData = separatedDataSheet.getDataRange().getValues();
    var assigneeData = assigneeConfigSheet.getDataRange().getValues();

    var separatedDataHeaders = separatedData[0];
    var assigneeHeaders = assigneeData[0];

    var assigneeHeaderMap = {};
    for (var i = 0; i < assigneeHeaders.length; i++) {
      assigneeHeaderMap[assigneeHeaders[i]] = i;
    }

    // Initialize result sheet
    if (slaConfigSheet) {
      slaConfigSheet.clear();
    } else {
      slaConfigSheet = ss.insertSheet(resultSheetName);
    }
    slaConfigSheet.setFrozenRows(1);
    slaConfigSheet.appendRow(["Category", "Sub Categories", "Department Name", "Form Name", "Employee Code", "Assignee Email", "Group", "Escalation 1", "Escalation Time 1", "Escalation 2", "Escalation Time 2", "Seprate Row", "Assignee Row", "Combination Status"]);

    var combinationsInSeparatedData = new Set();
    var combinationsInAssignee = new Set();
    var slaData = []; // For batch processing

    // Process separated data
    for (var i = 1; i < separatedData.length; i++) {
      var category = separatedData[i][separatedDataHeaders.indexOf("Category")];
      var subcategory = separatedData[i][separatedDataHeaders.indexOf("Sub Categories")];
      var Department = separatedData[i][separatedDataHeaders.indexOf("Department")];
      var employeeCode = separatedData[i][separatedDataHeaders.indexOf("Extra Field 1")];
 var FormName = separatedData[i][separatedDataHeaders.indexOf("Form Name")];
     
//      var combinationKey = useEmployeeCode ? (category + subcategory + employeeCode) : (category + subcategory);
  //    combinationsInSeparatedData.add(combinationKey);

    var combinationKey = useEmployeeCode ? (category + subcategory + employeeCode) : (category + subcategory);
    combinationsInSeparatedData.add(combinationKey);

      var matchFound = false;
      for (var j = 1; j < assigneeData.length; j++) {
        var assigneeCategory = assigneeData[j][assigneeHeaderMap["Category"]];
        var assigneeSubcategory = assigneeData[j][assigneeHeaderMap["Sub Categories"]];
        var assigneeEmployeeCode = assigneeData[j][assigneeHeaderMap["Extra Field 1"]];

        if (
          (assigneeCategory === category && (assigneeSubcategory === subcategory || assigneeSubcategory === "")) && // Match category and optionally subcategory
          (useEmployeeCode ? assigneeEmployeeCode === employeeCode : true)
        ) {
          var rowData = [];
          rowData.push(category, subcategory, Department,FormName, employeeCode);
          rowData.push(assigneeData[j][assigneeHeaderMap["Assignee Email"]]);
          rowData.push(assigneeData[j][assigneeHeaderMap["Group"]]);
          rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 1"]]);
          rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 1"]]);
          rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 2"]]);
          rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 2"]]);
          rowData.push(i + 1);
          rowData.push(j + 1);
          rowData.push(assigneeSubcategory === "" ? "Based on Category from Assignee Config" : "Found in both files");
          slaData.push(rowData);
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
        rowData.push(i + 1);
        rowData.push("N/A");
        rowData.push("Not found in assignee config");
        slaData.push(rowData);
      }

      if (slaData.length >= 100) {
        slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, slaData.length, slaData[0].length).setValues(slaData);
        slaData = [];
      }
    }

    if (slaData.length > 0) {
      slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, slaData.length, slaData[0].length).setValues(slaData);
    }

    // Check for extra records in Assignee config
    var extraRecordsInAssignee = [];
    for (var j = 1; j < assigneeData.length; j++) {
      var assigneeCategory = assigneeData[j][assigneeHeaderMap["Category"]];
      var assigneeSubcategory = assigneeData[j][assigneeHeaderMap["Sub Categories"]];
      var assigneeEmployeeCode =  assigneeData[j][assigneeHeaderMap["Extra Field 1"]];

      var matchFound = false;

      for (var i = 1; i < separatedData.length; i++) {
        var separatedCategory = separatedData[i][separatedDataHeaders.indexOf("Category")];
        var separatedSubcategory = separatedData[i][separatedDataHeaders.indexOf("Sub Categories")];
        var separatedEmployeeCode = separatedData[i][separatedDataHeaders.indexOf("Extra Field 1")];

        //var combinationKey = useEmployeeCode ? (separatedCategory + separatedSubcategory + separatedEmployeeCode) : (separatedCategory + separatedSubcategory);
    var combinationKey = assigneeCategory + assigneeSubcategory + (useEmployeeCode ? assigneeEmployeeCode : '');
    
        if (
          (assigneeCategory === separatedCategory && (assigneeSubcategory === separatedSubcategory || assigneeSubcategory === "")) && // Match category and optionally subcategory
          (useEmployeeCode ? assigneeEmployeeCode === separatedEmployeeCode : true)
        ) {
          matchFound = true;
          break;
        }
      }

      if (!matchFound) {
        var rowData = [];
        rowData.push(assigneeCategory, assigneeSubcategory, "N/A", "N/A", assigneeEmployeeCode);
        rowData.push(assigneeData[j][assigneeHeaderMap["Assignee Email"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Group"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 1"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 1"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation 2"]]);
        rowData.push(assigneeData[j][assigneeHeaderMap["Escalation Time 2"]]);
        rowData.push("N/A");
        rowData.push(j + 1);
        rowData.push("Not available in separated data");
        extraRecordsInAssignee.push(rowData);
      }
    }

    if (extraRecordsInAssignee.length > 0) {
      slaConfigSheet.getRange(slaConfigSheet.getLastRow() + 1, 1, extraRecordsInAssignee.length, extraRecordsInAssignee[0].length).setValues(extraRecordsInAssignee);
    }

    showProgressToast(SpreadsheetApp.getActiveSpreadsheet(), "Processing completed for: " + selectedOption);
    Logger.log("Processing completed for: " + selectedOption);
  } catch (error) {
    Logger.log("Error in createSLAConfig_Optimized_subcat_optional: " + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Error in creating SLA config: " + error.message, '⚠️ Error', 5);
  }
}

	function showProgressToast(ss, message) {
	  try {
		var ss = SpreadsheetApp.getActiveSpreadsheet();
		ss.toast(message, 'Progress', 5); // Display for 5 seconds
		SpreadsheetApp.flush(); // Ensure the UI updates are pushed out immediately
	  } catch (error) {
		Logger.log("Error in showProgressToast: " + error.message);
	  }
	}