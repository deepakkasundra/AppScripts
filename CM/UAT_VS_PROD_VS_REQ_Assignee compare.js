// Inputhset , compare sheet , param status, header
function UAT_vs_PROD_Compare_WithParam() {
  UAT_vs_Comparison_Compare("PROD Assignee Config", "UAT Assignee Config",true);
}

function UAT_vs_PROD_Compare_WithoutParam() {
  UAT_vs_Comparison_Compare("PROD Assignee Config","UAT Assignee Config",false);
}

function UAT_vs_Requirement_Compare_WithParam() {
  UAT_vs_Comparison_Compare("Requirement Assignee Config","UAT Assignee Config",true);
}

function UAT_vs_Requirement_Compare_WithoutParam() {
  UAT_vs_Comparison_Compare("Requirement Assignee Config","UAT Assignee Config",false);
}

// New function for PROD vs Requirement with Param
function PROD_vs_Requirement_Compare_WithParam() {
  UAT_vs_Comparison_Compare("PROD Assignee Config", "Requirement Assignee Config", true);
}

// New function for PROD vs Requirement without Param
function PROD_vs_Requirement_Compare_WithoutParam() {
  UAT_vs_Comparison_Compare("PROD Assignee Config", "Requirement Assignee Config", false);
}




function UAT_vs_Comparison_Compare(sheetname,comparesheetname,includeParam) {
  

// Extract the first word (prefix) from both sheet names
var comparesheetPrefix = comparesheetname.split(' ')[0];  // Extract first word of comparesheetname (e.g., "Requirement")
var basesheetPrefix = sheetname.split(' ')[0];  // Extract first word of sheetname (e.g., "PROD")

// Function to abbreviate the prefix if its length is greater than 4 characters
function abbreviatePrefix(prefix) {
  if (prefix.length > 4) {
    return prefix.substring(0, 3);  // Take the first three characters of the prefix
  }
  return prefix;
}

// Apply the abbreviation function to both prefixes
comparesheetPrefix = abbreviatePrefix(comparesheetPrefix);
basesheetPrefix = abbreviatePrefix(basesheetPrefix);

// Log the updated prefixes
Logger.log("Compare sheet prefix: " + comparesheetPrefix);
Logger.log("Base sheet prefix: " + basesheetPrefix);

   Logger.log("base Sheet name is " + sheetname);
   Logger.log("Compare sheet name " + comparesheetname);
   Logger.log("Param status is " + includeParam);
 // Logger.log("Column name is " + PROD_column_name);


  if (includeParam) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Running with additional param', 'Execution Status', 5);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('Running without additional param', 'Execution Status', 5);
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var uatSheet = spreadsheet.getSheetByName(comparesheetname);
  var prodSheet = spreadsheet.getSheetByName(sheetname);
  var consolidatedSheet = spreadsheet.getSheetByName("Comparison of " + sheetname +" vs " + comparesheetname);
  Logger.log("Start reading");
  var uatData = uatSheet.getDataRange().getValues();
  var prodData = prodSheet.getDataRange().getValues();
  Logger.log("Complete reading");
  SpreadsheetApp.getActiveSpreadsheet().toast('File Reading complete', 'Execution Status', 5);

  var requiredHeaders = ['Department', 'Category', 'Sub Categories',  'Assignee Email'];
  var missingHeaders = [];
  for (var header of requiredHeaders) {
    if (uatData[0].indexOf(header) === -1) {
      missingHeaders.push('UAT File: ' + header);
    }
    if (prodData[0].indexOf(header) === -1) {
      missingHeaders.push('Requirement File: ' + header);
    }
  }

  if (missingHeaders.length > 0) {
    var missingHeadersString = missingHeaders.join(', ');
    SpreadsheetApp.getActiveSpreadsheet().toast('The following required column headers are missing: ' + missingHeadersString, '⚠️ Warning', 10);
    return;
  }

  if (!consolidatedSheet) {
    consolidatedSheet = spreadsheet.insertSheet("Comparison of " + sheetname +" vs " + comparesheetname);
  Logger.log("output sheet created " + consolidatedSheet.getName());
  }

  Logger.log("Comparing started");
SpreadsheetApp.getActiveSpreadsheet().toast('Initiating Comparison', 'Execution Status', 5);

  consolidatedSheet.clear();
  consolidatedSheet.setFrozenRows(1);
  consolidatedSheet.setFrozenColumns(2);

var headers = [comparesheetPrefix + " Row", basesheetPrefix +" Row", "Name", "Department", "Category", "Sub Categories", comparesheetPrefix + " additional Param", basesheetPrefix +" additional Param", "Additional Param Status", comparesheetPrefix + " Assignee Email", basesheetPrefix +" Assignee Email", "Assignee Status", comparesheetPrefix + " CC Email", basesheetPrefix +" CC Email", "CC Email Status", comparesheetPrefix + " CC User Param", basesheetPrefix +" CC User Param", "CC User Param Status", comparesheetPrefix + " Comment Approver Email", basesheetPrefix +" Comment Approver Email", "Comment Approver Email Status",comparesheetPrefix + " Comment Approver Group",basesheetPrefix +" Comment Approver Group","Comment Approver Group Status", comparesheetPrefix + " Group", basesheetPrefix +" Group", "Group Status", comparesheetPrefix + " Assignee Email Param", basesheetPrefix +" Assignee Email Param", "Assignee Email Param Status", comparesheetPrefix + " Escalation 1", basesheetPrefix +" Escalation 1", "Escalation 1 Status",comparesheetPrefix + " Escalation Group 1", basesheetPrefix +" Escalation Group 1", "Escalation Group 1 Status",comparesheetPrefix + " Escalation param 1", basesheetPrefix +" Escalation param 1", "Escalation Param 1 Status",  comparesheetPrefix + " Escalation Time Urgent 1", basesheetPrefix +" Escalation Time Urgent 1", "Escalation Time 1 Status", comparesheetPrefix + " Escalation 2", basesheetPrefix +" Escalation 2", "Escalation 2 Status",comparesheetPrefix + " Escalation Group 2", basesheetPrefix +" Escalation Group 2", "Escalation Group 2 Status",comparesheetPrefix + " Escalation param 2", basesheetPrefix +" Escalation param 2", "Escalation Param 2 Status",  comparesheetPrefix + " Escalation Time Urgent 2", basesheetPrefix +" Escalation Time Urgent 2", "Escalation Time 2 Status",comparesheetPrefix + " Escalation 3", basesheetPrefix +" Escalation 3", "Escalation 3 Status",comparesheetPrefix + " Escalation Group 3", basesheetPrefix +" Escalation Group 3", "Escalation Group 3 Status",comparesheetPrefix + " Escalation param 3", basesheetPrefix +" Escalation param 3", "Escalation Param 3 Status",  comparesheetPrefix + " Escalation Time Urgent 3", basesheetPrefix +" Escalation Time Urgent 3", "Escalation Time 3 Status",comparesheetPrefix + " Escalation 4", basesheetPrefix +" Escalation 4", "Escalation 4 Status",comparesheetPrefix + " Escalation Group 4", basesheetPrefix +" Escalation Group 4", "Escalation Group 4 Status",comparesheetPrefix + " Escalation param 4", basesheetPrefix +" Escalation param 4", "Escalation Param 4 Status",  comparesheetPrefix + " Escalation Time Urgent 4", basesheetPrefix +" Escalation Time Urgent 4", "Escalation Time 4 Status", comparesheetPrefix + " Escalation 5", basesheetPrefix +" Escalation 5", "Escalation 5 Status",comparesheetPrefix + " Escalation Group 5", basesheetPrefix +" Escalation Group 5", "Escalation Group 5 Status",comparesheetPrefix + " Escalation param 5", basesheetPrefix +" Escalation param 5", "Escalation Param 5 Status",  comparesheetPrefix + " Escalation Time Urgent 5", basesheetPrefix +" Escalation Time Urgent 5", "Escalation Time 5 Status", comparesheetPrefix + " sorted Param", basesheetPrefix +" sorted Param", "Sorted Param Status", "Combination Status"]; 


//var headers = ["UAT Row", PROD_column_name +" Row", "Name", "Department", "Category", "Sub Categories", "UAT additional Param", PROD_column_name +". additional Param", "Additional Param Status", "UAT Assignee Email", PROD_column_name +". Assignee Email", "Assignee Status", "UAT CC Email", PROD_column_name +". CC Email", "CC Email Status", "UAT CC User Param", PROD_column_name +". CC User Param", "CC User Param Status", "UAT Comment Approver Email", PROD_column_name +". Comment Approver Email", "Comment Approver Email Status","UAT Comment Approver Group",PROD_column_name +". Comment Approver Group","Comment Approver Group Status", "UAT Group", PROD_column_name +". Group", "Group Status", "UAT Assignee Email Param", PROD_column_name +". Assignee Email Param", "Assignee Email Param Status", "UAT Escalation 1", PROD_column_name +". Escalation 1", "Escalation 1 Status","UAT Escalation Group 1", PROD_column_name +". Escalation Group 1", "Escalation Group 1 Status","UAT Escalation param 1", PROD_column_name +". Escalation param 1", "Escalation Param 1 Status",  "UAT Escalation Time Urgent 1", PROD_column_name +". Escalation Time Urgent 1", "Escalation Time 1 Status", "UAT Escalation 2", PROD_column_name +". Escalation 2", "Escalation 2 Status","UAT Escalation Group 2", PROD_column_name +". Escalation Group 2", "Escalation Group 2 Status","UAT Escalation param 2", PROD_column_name +". Escalation param 2", "Escalation Param 2 Status",  "UAT Escalation Time Urgent 2", PROD_column_name +". Escalation Time Urgent 2", "Escalation Time 2 Status","UAT Escalation 3", PROD_column_name +". Escalation 3", "Escalation 3 Status","UAT Escalation Group 3", PROD_column_name +". Escalation Group 3", "Escalation Group 3 Status","UAT Escalation param 3", PROD_column_name +". Escalation param 3", "Escalation Param 3 Status",  "UAT Escalation Time Urgent 3", PROD_column_name +". Escalation Time Urgent 3", "Escalation Time 3 Status","UAT Escalation 4", PROD_column_name +". Escalation 4", "Escalation 4 Status","UAT Escalation Group 4", PROD_column_name +". Escalation Group 4", "Escalation Group 4 Status","UAT Escalation param 4", PROD_column_name +". Escalation param 4", "Escalation Param 4 Status",  "UAT Escalation Time Urgent 4", PROD_column_name +". Escalation Time Urgent 4", "Escalation Time 4 Status", "UAT Escalation 5", PROD_column_name +". Escalation 5", "Escalation 5 Status","UAT Escalation Group 5", PROD_column_name +". Escalation Group 5", "Escalation Group 5 Status","UAT Escalation param 5", PROD_column_name +". Escalation param 5", "Escalation Param 5 Status",  "UAT Escalation Time Urgent 5", PROD_column_name +". Escalation Time Urgent 5", "Escalation Time 5 Status", "UAT sorted Param", PROD_column_name +" sorted Param", "Sorted Param Status", "Combination Status"]; 

spreadsheet.setActiveSheet(consolidatedSheet);
  consolidatedSheet.appendRow(headers);

  var uatHeaderMap = mapHeaders(uatData[0]);
  var prodHeaderMap = mapHeaders(prodData[0]);
  var processedCombinations = [];

  var batchSize = 1000; // Adjust batch size as needed
  var rowsToAppend = [];

  // Process UAT data for combination 
  for (var i = 1; i < uatData.length; i++) {
    var uatRow = uatData[i];
    var dept = uatRow[uatHeaderMap["Department"]];
    var category = uatRow[uatHeaderMap["Category"]];
    var subcategory = uatRow[uatHeaderMap["Sub Categories"]];
    var param = uatRow[uatHeaderMap["Execution Rule 1"]];
    var identifier = dept + category + subcategory;

    // Setting Identifier as per selection

      if (includeParam) {
      identifier += param;
    }
    
     
    // Logger.log(identifier)
    var prodRow = prodData.find(function(prodRow) {
      return (
        prodRow[prodHeaderMap["Department"]] === dept &&
        prodRow[prodHeaderMap["Category"]] === category &&
        prodRow[prodHeaderMap["Sub Categories"]] === subcategory && 
         (!includeParam || prodRow[prodHeaderMap["Execution Rule 1"]] === param)
     
      );
    });

    var uatExecutionRule = uatRow ? uatRow[uatHeaderMap["Execution Rule 1"] || ""] : "";
    var prodParam = prodRow ? prodRow[prodHeaderMap["Execution Rule 1"] || ""] : "";
    var ruleValidation = (uatExecutionRule === "" && prodParam === "") ? "NA" : (uatExecutionRule === prodParam) ? "Match" : "Not Match";

    // var processedUatdata = uatRow ? processRule(uatRow[uatHeaderMap["Execution Rule 1"] || ""]) : "";
    // var processedPRODdata = prodRow ? processRule(prodRow[prodHeaderMap["Execution Rule 1"] || ""]) : "";

var processedUatdata = uatRow ? uatRow[uatHeaderMap["Execution Rule 1"] || ""] : "";
var processedPRODdata = prodRow ? prodRow[prodHeaderMap["Execution Rule 1"] || ""] : "";

    var sorted_data_ruleValidation = (processedUatdata === "" && processedPRODdata === "") ? "NA" : (processedUatdata === processedPRODdata) ? "Match" : "Not Match";

 // var assigneeEmailUAT = uatRow[uatHeaderMap["Assignee Email"]];
//  var assigneeEmailPROD = prodRow ? prodRow[prodHeaderMap["Assignee Email"]] : "";
  
    var row = [
      comparesheetPrefix +" Row " + (i + 1),
      prodRow ? basesheetPrefix + " Row " + (prodData.indexOf(prodRow) + 1) : "",
      uatRow[uatHeaderMap["Name"]],
      dept,
      category,
      subcategory,
      uatExecutionRule,
      prodParam,
      ruleValidation,
      uatRow[uatHeaderMap["Assignee Email"]],
      prodRow ? prodRow[prodHeaderMap["Assignee Email"]] : "",
     
uatHeaderMap["Assignee Email"] !== undefined && prodHeaderMap["Assignee Email"] !== undefined ? 
    (uatRow[uatHeaderMap["Assignee Email"]].toLowerCase() === (prodRow && prodRow[prodHeaderMap["Assignee Email"]] !== undefined ? prodRow[prodHeaderMap["Assignee Email"]].toLowerCase() : "") ? 
        (uatRow[uatHeaderMap["Assignee Email"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

// Normalized Email ID even if it not in sequence
// uatHeaderMap["Assignee Email"] !== undefined && prodHeaderMap["Assignee Email"] !== undefined ?
//     (normalizeEmails(uatRow[uatHeaderMap["Assignee Email"]]) === (prodRow && prodRow[prodHeaderMap["Assignee Email"]] !== undefined ? normalizeEmails(prodRow[prodHeaderMap["Assignee Email"]]) : "") ?
//         (uatRow[uatHeaderMap["Assignee Email"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheets",

     // CC Email Status 
      uatRow[uatHeaderMap["CC Email"]],
      prodRow ? prodRow[prodHeaderMap["CC Email"]] : "",
//      uatRow[uatHeaderMap["CC Email"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["CC Email"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["CC Email"]] !== "" ? "Match" : "NA") : "Not Match", 
        uatHeaderMap["CC Email"] !== undefined && prodHeaderMap["CC Email"] !== undefined ? 
    (uatRow[uatHeaderMap["CC Email"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["CC Email"]].toLowerCase() : "") ? 
        (uatRow[uatHeaderMap["CC Email"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//CC User Param
      uatRow[uatHeaderMap["CC User Param"]],
      prodRow ? prodRow[prodHeaderMap["CC User Param"]] : "",
uatHeaderMap["CC User Param"] !== undefined && prodHeaderMap["CC User Param"] !== undefined ? 
    (uatRow[uatHeaderMap["CC User Param"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["CC User Param"]].toLowerCase() : "") ? 
        (uatRow[uatHeaderMap["CC User Param"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",
// Comment approver group
      uatRow[uatHeaderMap["Comment Approver Email"]],
      prodRow ? prodRow[prodHeaderMap["Comment Approver Email"]] : "",
uatHeaderMap["Comment Approver Email"] !== undefined && prodHeaderMap["Comment Approver Email"] !== undefined ? 
    (uatRow[uatHeaderMap["Comment Approver Email"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Comment Approver Email"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Comment Approver Email"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",
//comment approver group
      uatRow[uatHeaderMap["Comment Approver Group"]],
      prodRow ? prodRow[prodHeaderMap["Comment Approver Group"]] : "",
uatHeaderMap["Comment Approver Group"] !== undefined && prodHeaderMap["Comment Approver Group"] !== undefined ? 
    (uatRow[uatHeaderMap["Comment Approver Group"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Comment Approver Group"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Comment Approver Group"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

// group
      uatRow[uatHeaderMap["Group"]],
      prodRow ? prodRow[prodHeaderMap["Group"]] : "",
uatHeaderMap["Group"] !== undefined && prodHeaderMap["Group"] !== undefined ? 
    (uatRow[uatHeaderMap["Group"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Group"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Group"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",
// Assignee Email Param
      uatRow[uatHeaderMap["Assignee Email Param"]],
      prodRow ? prodRow[prodHeaderMap["Assignee Email Param"]] : "",
      uatHeaderMap["Assignee Email Param"] !== undefined && prodHeaderMap["Assignee Email Param"] !== undefined ? 
    (uatRow[uatHeaderMap["Assignee Email Param"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Assignee Email Param"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Assignee Email Param"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

// Escalation 1
      uatRow[uatHeaderMap["Escalation 1"]],
      prodRow ? prodRow[prodHeaderMap["Escalation 1"]] : "",
   
uatHeaderMap["Escalation 1"] !== undefined && prodHeaderMap["Escalation 1"] !== undefined ? 
    (String(uatRow[uatHeaderMap["Escalation 1"]] || "").toLowerCase() === (prodRow ? String(prodRow[prodHeaderMap["Escalation 1"]] || "").toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation 1"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Group 1
uatRow[uatHeaderMap["Escalation Group 1"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Group 1"]] : "",
      uatHeaderMap["Escalation Group 1"] !== undefined && prodHeaderMap["Escalation Group 1"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Group 1"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Group 1"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Group 1"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Param 1
      uatRow[uatHeaderMap["Escalation Param 1"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Param 1"]] : "",
      uatHeaderMap["Escalation Param 1"] !== undefined && prodHeaderMap["Escalation Param 1"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Param 1"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Param 1"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Param 1"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Time Urgent 1
      uatRow[uatHeaderMap["Escalation Time Urgent 1"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 1"]] : "",
      uatHeaderMap["Escalation Time Urgent 1"] !== undefined && prodHeaderMap["Escalation Time Urgent 1"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Time Urgent 1"]] === (prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 1"]] : "") ? (uatRow[uatHeaderMap["Escalation Time Urgent 1"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

// Escalation 2
      uatRow[uatHeaderMap["Escalation 2"]],
      prodRow ? prodRow[prodHeaderMap["Escalation 2"]] : "",
   
uatHeaderMap["Escalation 2"] !== undefined && prodHeaderMap["Escalation 2"] !== undefined ? 
    (String(uatRow[uatHeaderMap["Escalation 2"]] || "").toLowerCase() === (prodRow ? String(prodRow[prodHeaderMap["Escalation 2"]] || "").toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation 2"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Group 2
uatRow[uatHeaderMap["Escalation Group 2"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Group 2"]] : "",
      uatHeaderMap["Escalation Group 2"] !== undefined && prodHeaderMap["Escalation Group 2"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Group 2"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Group 2"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Group 2"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Param 2
      uatRow[uatHeaderMap["Escalation Param 2"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Param 2"]] : "",
      uatHeaderMap["Escalation Param 2"] !== undefined && prodHeaderMap["Escalation Param 2"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Param 2"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Param 2"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Param 2"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Time Urgent 2
      uatRow[uatHeaderMap["Escalation Time Urgent 2"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 2"]] : "",
      uatHeaderMap["Escalation Time Urgent 2"] !== undefined && prodHeaderMap["Escalation Time Urgent 2"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Time Urgent 2"]] === (prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 2"]] : "") ? (uatRow[uatHeaderMap["Escalation Time Urgent 2"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

// Escalation 3
      uatRow[uatHeaderMap["Escalation 3"]],
      prodRow ? prodRow[prodHeaderMap["Escalation 3"]] : "",
   
uatHeaderMap["Escalation 3"] !== undefined && prodHeaderMap["Escalation 3"] !== undefined ? 
    (String(uatRow[uatHeaderMap["Escalation 3"]] || "").toLowerCase() === (prodRow ? String(prodRow[prodHeaderMap["Escalation 3"]] || "").toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation 3"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Group 3
uatRow[uatHeaderMap["Escalation Group 3"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Group 3"]] : "",
      uatHeaderMap["Escalation Group 3"] !== undefined && prodHeaderMap["Escalation Group 3"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Group 3"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Group 3"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Group 3"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Param 3
      uatRow[uatHeaderMap["Escalation Param 3"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Param 3"]] : "",
      uatHeaderMap["Escalation Param 3"] !== undefined && prodHeaderMap["Escalation Param 3"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Param 3"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Param 3"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Param 3"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",
//Escalation Time Urgent 3
      uatRow[uatHeaderMap["Escalation Time Urgent 3"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 3"]] : "",
      uatHeaderMap["Escalation Time Urgent 3"] !== undefined && prodHeaderMap["Escalation Time Urgent 3"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Time Urgent 3"]] === (prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 3"]] : "") ? (uatRow[uatHeaderMap["Escalation Time Urgent 3"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

// Escalation 4
      uatRow[uatHeaderMap["Escalation 4"]],
      prodRow ? prodRow[prodHeaderMap["Escalation 4"]] : "",
   
uatHeaderMap["Escalation 4"] !== undefined && prodHeaderMap["Escalation 4"] !== undefined ? 
    (String(uatRow[uatHeaderMap["Escalation 4"]] || "").toLowerCase() === (prodRow ? String(prodRow[prodHeaderMap["Escalation 4"]] || "").toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation 4"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Group 4
uatRow[uatHeaderMap["Escalation Group 4"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Group 4"]] : "",
      uatHeaderMap["Escalation Group 4"] !== undefined && prodHeaderMap["Escalation Group 4"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Group 4"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Group 4"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Group 4"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Param 4
      uatRow[uatHeaderMap["Escalation Param 4"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Param 4"]] : "",
      uatHeaderMap["Escalation Param 4"] !== undefined && prodHeaderMap["Escalation Param 4"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Param 4"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Param 4"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Param 4"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Time Urgent 4
      uatRow[uatHeaderMap["Escalation Time Urgent 4"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 4"]] : "",
      uatHeaderMap["Escalation Time Urgent 4"] !== undefined && prodHeaderMap["Escalation Time Urgent 4"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Time Urgent 4"]] === (prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 4"]] : "") ? (uatRow[uatHeaderMap["Escalation Time Urgent 4"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",


// Escalation 5
      uatRow[uatHeaderMap["Escalation 5"]],
      prodRow ? prodRow[prodHeaderMap["Escalation 5"]] : "",
   
uatHeaderMap["Escalation 5"] !== undefined && prodHeaderMap["Escalation 5"] !== undefined ? 
    (String(uatRow[uatHeaderMap["Escalation 5"]] || "").toLowerCase() === (prodRow ? String(prodRow[prodHeaderMap["Escalation 5"]] || "").toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation 5"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Group 5
uatRow[uatHeaderMap["Escalation Group 5"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Group 5"]] : "",
      uatHeaderMap["Escalation Group 5"] !== undefined && prodHeaderMap["Escalation Group 5"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Group 5"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Group 5"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Group 5"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Param 5
      uatRow[uatHeaderMap["Escalation Param 5"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Param 5"]] : "",
      uatHeaderMap["Escalation Param 5"] !== undefined && prodHeaderMap["Escalation Param 5"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Param 5"]].toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Escalation Param 5"]].toLowerCase() : "") ? (uatRow[uatHeaderMap["Escalation Param 5"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",

//Escalation Time Urgent 5
      uatRow[uatHeaderMap["Escalation Time Urgent 5"]],
      prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 5"]] : "",
      uatHeaderMap["Escalation Time Urgent 5"] !== undefined && prodHeaderMap["Escalation Time Urgent 5"] !== undefined ? 
    (uatRow[uatHeaderMap["Escalation Time Urgent 5"]] === (prodRow ? prodRow[prodHeaderMap["Escalation Time Urgent 5"]] : "") ? (uatRow[uatHeaderMap["Escalation Time Urgent 5"]] !== "" ? "Match" : "NA") : "Not Match") : "Column Not Available in one of the sheet",
 
      processedUatdata,
      processedPRODdata,
      sorted_data_ruleValidation,
      prodRow ? "Available in Both" : "Not in " + basesheetPrefix
    ];

    rowsToAppend.push(row);
    processedCombinations.push(identifier);

    // Batch processing: Append rows when batch size is reached
    if (i % batchSize === 0 || i === uatData.length - 1) {
      Logger.log("Processing Batch: " + Math.ceil(i / batchSize)); // Log batch number
      consolidatedSheet.getRange(consolidatedSheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
      rowsToAppend = []; // Reset rowsToAppend array for the next batch
    }

    // Clear data to avoid same data in next process
    uatRow = null;
    prodRow = null;
    uatExecutionRule = null;
    prodParam = null;
    processedUatdata = null;
    processedPRODdata = null;
    
  }

  // Process PROD data
  for (var i = 1; i < prodData.length; i++) {
    
    var prodRow = prodData[i];
    var dept = prodRow[prodHeaderMap["Department"]];
    var category = prodRow[prodHeaderMap["Category"]];
    var subcategory = prodRow[prodHeaderMap["Sub Categories"]];
    var param = prodRow[prodHeaderMap["Execution Rule 1"]];
    var identifier = dept + category + subcategory;

      if (includeParam) {
      identifier += param;
    }

   // Logger.log("PROD data process" + identifier);

    var uatExecutionRule = uatRow ? uatRow[uatHeaderMap["Execution Rule 1"] || ""] : "";
    var prodParam = prodRow ? prodRow[prodHeaderMap["Execution Rule 1"] || ""] : "";
    var ruleValidation = (uatExecutionRule === "" && prodParam === "") ? "NA" : (uatExecutionRule === prodParam) ? "Match" : "Not Match";

    // var processedUatdata = uatRow ? processRule(uatRow[uatHeaderMap["Execution Rule 1"] || ""]) : "";
    // var processedPRODdata = prodRow ? processRule(prodRow[prodHeaderMap["Execution Rule 1"] || ""]) : "";
    
    var processedUatdata = uatRow ? uatRow[uatHeaderMap["Execution Rule 1"] || ""] : "";
    var processedPRODdata = prodRow ? prodRow[prodHeaderMap["Execution Rule 1"] || ""] : "";
    
    var sorted_data_ruleValidation = (processedUatdata === "" && processedPRODdata === "") ? "NA" : (processedUatdata === processedPRODdata) ? "Match" : "Not Match";


    if (!processedCombinations.includes(identifier)) {
      var uatExecutionRule = "";
      var prodParam = prodRow[prodHeaderMap["Execution Rule 1"]];
      var ruleValidation = "Not in UAT";

      var row = [
        "",
        basesheetPrefix + ". Row " + (i + 1),
        prodRow[prodHeaderMap["Name"]],
        dept,
        category,
        subcategory,
        uatExecutionRule,
        prodParam,
        ruleValidation,

        "",
        prodRow[prodHeaderMap["Assignee Email"]],
     // validate Assignee Email
     ((uatRow && uatHeaderMap && uatHeaderMap["Assignee Email"] !== undefined ? uatRow[uatHeaderMap["Assignee Email"]] : "") === "" && 
    (prodRow ? prodRow[prodHeaderMap["Assignee Email"]] : "") === ""  ) ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Assignee Email"] !== undefined ? uatRow[uatHeaderMap["Assignee Email"]] : "").toLowerCase() === (prodRow ? prodRow[prodHeaderMap["Assignee Email"]] : "").toLowerCase()) ? "Match" : "Not Match",
      

        "",
        prodRow[prodHeaderMap["CC Email"]],
   (!uatHeaderMap || uatHeaderMap["CC Email"] === undefined || !prodHeaderMap || prodHeaderMap["CC Email"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["CC Email"] !== undefined ? uatRow[uatHeaderMap["CC Email"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["CC Email"] !== undefined ? prodRow[prodHeaderMap["CC Email"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["CC Email"] !== undefined ? uatRow[uatHeaderMap["CC Email"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["CC Email"] !== undefined ? prodRow[prodHeaderMap["CC Email"]] : "").toLowerCase()) ? "Match" : "Not Match",
        
        "",
       prodRow[prodHeaderMap["CC User Param"]],
        (!uatHeaderMap || uatHeaderMap["CC User Param"] === undefined || !prodHeaderMap || prodHeaderMap["CC User Param"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["CC User Param"] !== undefined ? uatRow[uatHeaderMap["CC User Param"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["CC User Param"] !== undefined ? prodRow[prodHeaderMap["CC User Param"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["CC User Param"] !== undefined ? uatRow[uatHeaderMap["CC User Param"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["CC User Param"] !== undefined ? prodRow[prodHeaderMap["CC User Param"]] : "").toLowerCase()) ? "Match" : "Not Match",

// Comment approver
        "",
    prodRow[prodHeaderMap["Comment Approver Email"]],
   
   (!uatHeaderMap || uatHeaderMap["Comment Approver Email"] === undefined || !prodHeaderMap || prodHeaderMap["Comment Approver Email"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Comment Approver Email"] !== undefined ? uatRow[uatHeaderMap["Comment Approver Email"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Comment Approver Email"] !== undefined ? prodRow[prodHeaderMap["Comment Approver Email"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Comment Approver Email"] !== undefined ? uatRow[uatHeaderMap["Comment Approver Email"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Comment Approver Email"] !== undefined ? prodRow[prodHeaderMap["Comment Approver Email"]] : "").toLowerCase()
) ? "Match" : "Not Match",

// Comment approver Group
        "",
    prodRow[prodHeaderMap["Comment Approver Group"]],
   (!uatHeaderMap || uatHeaderMap["Comment Approver Group"] === undefined || !prodHeaderMap || prodHeaderMap["Comment Approver Group"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Comment Approver Group"] !== undefined ? uatRow[uatHeaderMap["Comment Approver Group"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Comment Approver Group"] !== undefined ? prodRow[prodHeaderMap["Comment Approver Group"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Comment Approver Group"] !== undefined ? uatRow[uatHeaderMap["Comment Approver Group"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Comment Approver Group"] !== undefined ? prodRow[prodHeaderMap["Comment Approver Group"]] : "").toLowerCase()
) ? "Match" : "Not Match",


        // Group
"",
    prodRow[prodHeaderMap["Group"]],
   (!uatHeaderMap || uatHeaderMap["Group"] === undefined || !prodHeaderMap || prodHeaderMap["Group"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Group"] !== undefined ? uatRow[uatHeaderMap["Group"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Group"] !== undefined ? prodRow[prodHeaderMap["Group"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Group"] !== undefined ? uatRow[uatHeaderMap["Group"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Group"] !== undefined ? prodRow[prodHeaderMap["Group"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Assignee Email Param
    "",
    prodRow[prodHeaderMap["Assignee Email Param"]],
   (!uatHeaderMap || uatHeaderMap["Assignee Email Param"] === undefined || !prodHeaderMap || prodHeaderMap["Assignee Email Param"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Assignee Email Param"] !== undefined ? uatRow[uatHeaderMap["Assignee Email Param"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Assignee Email Param"] !== undefined ? prodRow[prodHeaderMap["Assignee Email Param"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Assignee Email Param"] !== undefined ? uatRow[uatHeaderMap["Assignee Email Param"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Assignee Email Param"] !== undefined ? prodRow[prodHeaderMap["Assignee Email Param"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escalation 1

        "",
        prodRow[prodHeaderMap["Escalation 1"]],
   (!uatHeaderMap || uatHeaderMap["Escalation 1"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation 1"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 1"] !== undefined ? uatRow[uatHeaderMap["Escalation 1"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation 1"] !== undefined ? prodRow[prodHeaderMap["Escalation 1"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 1"] !== undefined ? uatRow[uatHeaderMap["Escalation 1"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation 1"] !== undefined ? prodRow[prodHeaderMap["Escalation 1"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escaltion 1 Group - Escalation Group 1
        "",
        prodRow[prodHeaderMap["Escalation Group 1"]],
   (!uatHeaderMap || uatHeaderMap["Escalation Group 1"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Group 1"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 1"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 1"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 1"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 1"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 1"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 1"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 1"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 1"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escalation Param 1
        "",
        prodRow[prodHeaderMap["Escalation Param 1"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Param 1"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Param 1"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 1"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 1"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 1"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 1"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 1"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 1"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 1"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 1"]] : "").toLowerCase()) ? "Match" : "Not Match",


        // Escalation Time Urgent 1
        "",
        prodRow[prodHeaderMap["Escalation Time Urgent 1"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Time Urgent 1"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Time Urgent 1"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 1"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 1"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 1"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 1"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 1"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 1"]] : "") === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 1"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 1"]] : "")) ? "Match" : "Not Match",

// Escalation 2

        "",
        prodRow[prodHeaderMap["Escalation 2"]],
   (!uatHeaderMap || uatHeaderMap["Escalation 2"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation 2"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 2"] !== undefined ? uatRow[uatHeaderMap["Escalation 2"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation 2"] !== undefined ? prodRow[prodHeaderMap["Escalation 2"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 2"] !== undefined ? uatRow[uatHeaderMap["Escalation 2"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation 2"] !== undefined ? prodRow[prodHeaderMap["Escalation 2"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escaltion 2 Group - Escalation Group 2
        "",
        prodRow[prodHeaderMap["Escalation Group 2"]],
   (!uatHeaderMap || uatHeaderMap["Escalation Group 2"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Group 2"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 2"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 2"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 2"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 2"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 2"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 2"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 2"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 2"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escalation Param 2
        "",
        prodRow[prodHeaderMap["Escalation Param 2"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Param 2"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Param 2"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 2"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 2"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 2"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 2"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 2"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 2"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 2"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 2"]] : "").toLowerCase()) ? "Match" : "Not Match",


        // Escalation Time Urgent 2
        "",
        prodRow[prodHeaderMap["Escalation Time Urgent 2"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Time Urgent 2"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Time Urgent 2"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 2"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 2"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 2"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 2"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 2"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 2"]] : "") === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 2"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 2"]] : "")) ? "Match" : "Not Match",

// Escalation 3

        "",
        prodRow[prodHeaderMap["Escalation 3"]],
   (!uatHeaderMap || uatHeaderMap["Escalation 3"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation 3"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 3"] !== undefined ? uatRow[uatHeaderMap["Escalation 3"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation 3"] !== undefined ? prodRow[prodHeaderMap["Escalation 3"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 3"] !== undefined ? uatRow[uatHeaderMap["Escalation 3"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation 3"] !== undefined ? prodRow[prodHeaderMap["Escalation 3"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escaltion 3 Group - Escalation Group 3
        "",
        prodRow[prodHeaderMap["Escalation Group 3"]],
   (!uatHeaderMap || uatHeaderMap["Escalation Group 3"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Group 3"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 3"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 3"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 3"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 3"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 3"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 3"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 3"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 3"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escalation Param 3
        "",
        prodRow[prodHeaderMap["Escalation Param 3"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Param 3"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Param 3"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 3"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 3"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 3"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 3"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 3"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 3"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 3"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 3"]] : "").toLowerCase()) ? "Match" : "Not Match",


        // Escalation Time Urgent 3
        "",
        prodRow[prodHeaderMap["Escalation Time Urgent 3"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Time Urgent 3"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Time Urgent 3"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 3"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 3"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 3"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 3"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 3"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 3"]] : "") === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 3"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 3"]] : "")) ? "Match" : "Not Match",

// Escalation 4

        "",
        prodRow[prodHeaderMap["Escalation 4"]],
   (!uatHeaderMap || uatHeaderMap["Escalation 4"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation 4"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 4"] !== undefined ? uatRow[uatHeaderMap["Escalation 4"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation 4"] !== undefined ? prodRow[prodHeaderMap["Escalation 4"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 4"] !== undefined ? uatRow[uatHeaderMap["Escalation 4"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation 4"] !== undefined ? prodRow[prodHeaderMap["Escalation 4"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escaltion 4 Group - Escalation Group 4
        "",
        prodRow[prodHeaderMap["Escalation Group 4"]],
   (!uatHeaderMap || uatHeaderMap["Escalation Group 4"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Group 4"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 4"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 4"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 4"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 4"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 4"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 4"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 4"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 4"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escalation Param 4
        "",
        prodRow[prodHeaderMap["Escalation Param 4"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Param 4"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Param 4"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 4"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 4"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 4"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 4"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 4"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 4"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 4"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 4"]] : "").toLowerCase()) ? "Match" : "Not Match",


        // Escalation Time Urgent 4
        "",
        prodRow[prodHeaderMap["Escalation Time Urgent 4"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Time Urgent 4"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Time Urgent 4"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 4"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 4"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 4"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 4"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 4"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 4"]] : "") === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 4"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 4"]] : "")) ? "Match" : "Not Match",

// Escalation 5

        "",
        prodRow[prodHeaderMap["Escalation 5"]],
   (!uatHeaderMap || uatHeaderMap["Escalation 5"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation 5"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 5"] !== undefined ? uatRow[uatHeaderMap["Escalation 5"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation 5"] !== undefined ? prodRow[prodHeaderMap["Escalation 5"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation 5"] !== undefined ? uatRow[uatHeaderMap["Escalation 5"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation 5"] !== undefined ? prodRow[prodHeaderMap["Escalation 5"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escaltion 5 Group - Escalation Group 5
        "",
        prodRow[prodHeaderMap["Escalation Group 5"]],
   (!uatHeaderMap || uatHeaderMap["Escalation Group 5"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Group 5"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 5"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 5"]] : "") === "" && 
    (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 5"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 5"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Group 5"] !== undefined ? uatRow[uatHeaderMap["Escalation Group 5"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Group 5"] !== undefined ? prodRow[prodHeaderMap["Escalation Group 5"]] : "").toLowerCase()
) ? "Match" : "Not Match",


// Escalation Param 5
        "",
        prodRow[prodHeaderMap["Escalation Param 5"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Param 5"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Param 5"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 5"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 5"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 5"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 5"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Param 5"] !== undefined ? uatRow[uatHeaderMap["Escalation Param 5"]] : "").toLowerCase() === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Param 5"] !== undefined ? prodRow[prodHeaderMap["Escalation Param 5"]] : "").toLowerCase()) ? "Match" : "Not Match",


        // Escalation Time Urgent 5
        "",
        prodRow[prodHeaderMap["Escalation Time Urgent 5"]],
        (!uatHeaderMap || uatHeaderMap["Escalation Time Urgent 5"] === undefined || !prodHeaderMap || prodHeaderMap["Escalation Time Urgent 5"] === undefined ) ? "Column Not Available in one of the sheet" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 5"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 5"]] : "") === "" && (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 5"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 5"]] : "") === "") ? "NA" : ((uatRow && uatHeaderMap && uatHeaderMap["Escalation Time Urgent 5"] !== undefined ? uatRow[uatHeaderMap["Escalation Time Urgent 5"]] : "") === (prodRow && prodHeaderMap && prodHeaderMap["Escalation Time Urgent 5"] !== undefined ? prodRow[prodHeaderMap["Escalation Time Urgent 5"]] : "")) ? "Match" : "Not Match",



        processedUatdata,
        processedPRODdata,
        sorted_data_ruleValidation,
        "Not in UAT"
      ];

      rowsToAppend.push(row);
      processedCombinations.push(identifier);
    }
  }

  // Append remaining rows if not appended already
  if (rowsToAppend.length > 0) {
    consolidatedSheet.getRange(consolidatedSheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }

  // Apply conditional formatting
  applyConditionalFormatting_result(consolidatedSheet, headers,comparesheetPrefix,basesheetPrefix);

  SpreadsheetApp.getActiveSpreadsheet().toast('Comparison report generated successfully.', '👍 Process completed ', 10);
}

function applyConditionalFormatting_result(sheet, headers,comparesheetPrefix,basesheetPrefix) {
  // Clear existing conditional formatting rules
  sheet.clearConditionalFormatRules();

  // ColumnA duplicate
  var ColumnA_duplicate = sheet.getRange(`${getColumnName(headers.indexOf(comparesheetPrefix+" Row")+1)}:${getColumnName(headers.indexOf(comparesheetPrefix+" Row")+1)}`);
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=COUNTIF(A:A, $A1)>1`)
    .setBackground("#f4cccc")
    .setRanges([ColumnA_duplicate])
    .build();

// Logger.log(PROD_column_name+" Row");

  // ColumnB duplicate
  var ColumnB_duplicate = sheet.getRange(`${getColumnName(headers.indexOf(basesheetPrefix+" Row")+1)}:${getColumnName(headers.indexOf(basesheetPrefix+" Row")+1)}`);
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=COUNTIF(B:B, $B1)>1`)
    .setBackground("#f4cccc")
    .setRanges([ColumnB_duplicate])
    .build();

  // Create a rule to highlight cells containing "Not Match"
  var lastrange = sheet.getLastColumn();
  var rangeAtoZ = sheet.getRange(1, 1, sheet.getLastRow(), lastrange);
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Not Match")
    .setBackground("#f4cccc")
    .setRanges([rangeAtoZ])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(rule1, rule2, rule3);
  sheet.setConditionalFormatRules(rules);
}

function mapHeaders(headers) {
  var headerMap = {};
  for (var i = 0; i < headers.length; i++) {
    headerMap[headers[i]] = i;
  }
  return headerMap;
}

// function processRule(rule) {
//   return rule
//     .replace(/country:/g, "")
//     .replace(/location:/g, "")
//     .replace(/\|\|/g, ";")
//     .split(';')
//     .filter(value => value.trim() !== '')
//     .sort()
//     .join(';');
// }

function getColumnName(columnIndex) {
  var dividend = columnIndex;
  var columnName = '';
  var modulo;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}



// Helper function to normalize email lists
function normalizeEmails(emailString) {
    if (!emailString) return ""; // Handle empty or undefined email strings
    return emailString
        .split(";") // Split by semicolon
        .map(email => email.trim().toLowerCase()) // Trim and convert to lowercase
        .sort() // Sort the emails alphabetically
        .join(";"); // Join them back into a string
}
