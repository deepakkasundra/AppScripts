function workLMDynamicFormula_NA()
{
try{
  setFormulacommon_NA("WorkLM");

}  catch (e) {
  handleError(e);
  }
}

function AutonomousDynamicFormula_NA()
{
  try{
  setFormulacommon_NA("Autonomous");

}  catch (e) {
  handleError(e);
  }
}

function setFormulacommon_NA(FormulaSelection) {
try{
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QA_Report'); // Replace 'YourSheetName' with the name of your sheet.

  unhideAllColumns_NA(sheet);

  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerToSearch = "Policy Name"; // Change this to the header you want to search for

  // Find the index of the header dynamically
  var columnIndex = headers.indexOf(headerToSearch) + 1; // Add 1 to make it 1-based index

  if (columnIndex === 0) {
    Logger.log(headerToSearch + " column not found.");
    return;
  }

  // Define the formula dynamically based on the header
  var formulas;
  if (FormulaSelection === "WorkLM"){
    // Hide Autonoums Header
     hideColumns_NA(sheet, headers, [
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

  formulas = {
    // Policy Status
    "Policy Testing Status": `=IFS(${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2=0,"WIP",${getColumnName_NA(headers.indexOf("Need to be Tested") + 1)}2>0,"WIP",${getColumnName_NA(headers.indexOf("% Response from GPT") + 1)}2="NA","WIP",${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2=${getColumnName_NA(headers.indexOf("Pass") + 1)}2+${getColumnName_NA(headers.indexOf("Fail") + 1)}2,"Done")`,

// Test Cases count
  
  //  "Test Cases Count":`=IF(${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2="",0,COUNTA(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!B2:B")))`,

// below formula updated

"Test Cases Count": `=IF(OR(${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2="", ISERROR(INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!A1"))), 0, COUNTA(INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!" & ADDRESS(2, MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)) & ":" & ADDRESS(ROWS(INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!A:A")), MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)))))`
,
 
  // Pass
    "Pass": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName_NA(headers.indexOf("Pass") + 1)}$1)`,
    // Fail
    "Fail": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName_NA(headers.indexOf("Fail") + 1)}$1)`,
    // Need to be Tested
    "Need to be Tested": `=${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2-${getColumnName_NA(headers.indexOf("Pass") + 1)}2-${getColumnName_NA(headers.indexOf("Fail") + 1)}2`,
   // Response Count from GPT
    "Response Count from GPT" : `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"GPT")`,
    // % Response from GPT
    "% Response from GPT" : `=if(${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2=0,"NA",(${getColumnName_NA(headers.indexOf("Response Count from GPT") + 1)}2/${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2))`, // removed *100)    
    // GPT Pass %    
    "GPT Pass %" : `=if(${getColumnName_NA(headers.indexOf("GPT Pass Count") + 1)}2=0,"NA",(${getColumnName_NA(headers.indexOf("GPT Pass Count") + 1)}2/${getColumnName_NA(headers.indexOf("Response Count from GPT") + 1)}2))`, // removed *100)
    
    
    // GPT Fail %
    "GPT Fail %" : `=if(${getColumnName_NA(headers.indexOf("GPT Fail Count") + 1)}2=0,"0",(${getColumnName_NA(headers.indexOf("GPT Fail Count") + 1)}2/${getColumnName_NA(headers.indexOf("Response Count from GPT") + 1)}2))`, // removed *100)
    // Fallback Count
    "Fallback Count" : `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"Fallback")`,
    // FallBack
    "Fallback %" : `=if(${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2=0,"NA",(${getColumnName_NA(headers.indexOf("Fallback Count") + 1)}2/${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2))`, // removed *100
    // response from search count
    "response from search count" : `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"Search")`,
    // % response from Search
    "% response from Search" : `=if(${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2=0,"NA",(${getColumnName_NA(headers.indexOf("response from search count") + 1)}2/${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2))`, // removed *100
    // GPT Pass Count
    "GPT Pass Count" : `=COUNTIFs(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName_NA(headers.indexOf("Pass") + 1)}$1,INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"GPT")`,
    // GPT Fail Count
    "GPT Fail Count" : `=COUNTIFs(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName_NA(headers.indexOf("Fail") + 1)}$1,INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"GPT")`
  
  };

sheet.getRange(`${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}:${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}`).clearFormat();
Logger.log("Clear Condition format "  + `${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}:${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}`);
sheet.getRange(`${getColumnName_NA(headers.indexOf("% Response from GPT")+1)}:${getColumnName_NA(headers.indexOf("% Response from GPT")+1)}`).clearFormat();
Logger.log("Clear Condition format " + `${getColumnName_NA(headers.indexOf("% Response from GPT")+1 )}:${getColumnName_NA(headers.indexOf("% Response from GPT")+1)}`);


  // Set the dynamic formulas in the cells below their respective headers
  for (var header in formulas) {
    var columnIndex = headers.indexOf(header) + 1; // Get the index of the header
    if (columnIndex > 0) {

      var formula = formulas[header];
      var cellBelowHeader = sheet.getRange(2, columnIndex);
      cellBelowHeader.setFormula(formula).setBorder(true, true, true, true, true, true, "SOLID", null);
      var cellBelowHeaderformat = sheet.getRange(1, columnIndex);
      // set formating 
      cellBelowHeaderformat.setFontWeight('bold').setBackground("#c9daf8").setVerticalAlignment("top").setBorder(true, true, true, true, true, true, "SOLID", null);
Logger.log("Processed "  + formula);      
      

      // Format the percentage columns
      if (["% Response from GPT", "GPT Pass %", "GPT Fail %", "Fallback %"].includes(header)) {
        sheet.getRange(2, columnIndex, sheet.getMaxRows() - 1).setNumberFormat("0.00%"); // Format as percentage
      }
      
    }
  }


var nottestedRNG = sheet.getRange(`${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}:${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}`);
var rule1 = SpreadsheetApp.newConditionalFormatRule()
.whenNumberGreaterThan(0)
    .setBackground("#f4cccc")
    .setRanges([nottestedRNG])
    .build();
    
// GPT Range condition Format
var GPTNARNG = sheet.getRange(`${getColumnName_NA(headers.indexOf("% Response from GPT")+1)}:${getColumnName_NA(headers.indexOf("% Response from GPT")+1)}`);
var rule2 = SpreadsheetApp.newConditionalFormatRule()
.whenTextContains("NA")
    .setBackground("#fce5cd")
    .setRanges([GPTNARNG])
    .build();


var rules = sheet.getConditionalFormatRules();
 rules.push(rule1,rule2);
sheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getActiveSpreadsheet().toast("WorkLM/GPT formula have been updated", "👍 Formula Updated", 5);
  colorHeaders_NA();
  }
  else if (FormulaSelection === "Autonomous")  
  {

    hideColumns_NA(sheet, headers,[
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


  formulas = {
    // Policy Status
    "Policy Testing Status": `=IFS(${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2=0,"WIP",${getColumnName_NA(headers.indexOf("Need to be Tested") + 1)}2>0,"WIP",${getColumnName_NA(headers.indexOf("% Response from GPT") + 1)}2="NA","WIP",${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2=${getColumnName_NA(headers.indexOf("Pass") + 1)}2+${getColumnName_NA(headers.indexOf("Fail") + 1)}2,"Done")`,

// Test Cases count
  //  "Test Cases Count":`=IF(${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2="",0,COUNTA(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!B2:B")))`,
// below formula updated
"Test Cases Count": `=IF(OR(${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2="", ISERROR(INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!A1"))), 0, COUNTA(INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!" & ADDRESS(2, MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)) & ":" & ADDRESS(ROWS(INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!A:A")), MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)))))`,
  // Pass
    "Pass": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName_NA(headers.indexOf("Pass") + 1)}$1)`,
    // Fail
    "Fail": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName_NA(headers.indexOf("Fail") + 1)}$1)`,
    // Need to be Tested
    "Need to be Tested": `=${getColumnName_NA(headers.indexOf("Test Cases Count") + 1)}2-${getColumnName_NA(headers.indexOf("Pass") + 1)}2-${getColumnName_NA(headers.indexOf("Fail") + 1)}2`,
    "Bot stuck": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),$${getColumnName_NA(headers.indexOf("Bot stuck") + 1)}$1)`,
    "Incorrect audience": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("Incorrect audience") + 1)}$1)`,

    "Incorrect response": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("Incorrect response") + 1)}$1)`,
  
    "hallucination issue": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("hallucination issue") + 1)}$1)`,

    "Policy search": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("Policy search") + 1)}$1)`,
    
    "No/Fallback response": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("No/Fallback response") + 1)}$1)`,
    
    "Partial response given": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("Partial response given") + 1)}$1)`,
    
    "Generic response": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("Generic response") + 1)}$1)`,
    
    "Incorrect language": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("Incorrect language") + 1)}$1)`,
    
    "Incorrect Highlighting": `=COUNTIF(INDIRECT("'"&${getColumnName_NA(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName_NA(headers.indexOf("Incorrect Highlighting") + 1)}$1)`

  };
  // Set the dynamic formulas in the cells below their respective headers
  for (var header in formulas) {
    var columnIndex = headers.indexOf(header) + 1; // Get the index of the header
    if (columnIndex > 0) {

      var formula = formulas[header];
      var cellBelowHeader = sheet.getRange(2, columnIndex);
      cellBelowHeader.setFormula(formula).setBorder(true, true, true, true, true, true, "SOLID", null);
      var cellBelowHeaderformat = sheet.getRange(1, columnIndex);
      // set formating 
      cellBelowHeaderformat.setFontWeight('bold').setBackground("#F4CCCC").setVerticalAlignment("top").setBorder(true, true, true, true, true, true, "SOLID", null);
Logger.log("Processed "  + formula);      
    
      
    }
  }

var nottestedRNG = sheet.getRange(`${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}:${getColumnName_NA(headers.indexOf("Need to be Tested")+1)}`);
var rule1 = SpreadsheetApp.newConditionalFormatRule()
.whenNumberGreaterThan(0)
    .setBackground("#f4cccc")
    .setRanges([nottestedRNG])
    .build();
    
// GPT Range condition Format
var GPTNARNG = sheet.getRange(`${getColumnName_NA(headers.indexOf("% Response from GPT")+1)}:${getColumnName_NA(headers.indexOf("% Response from GPT")+1)}`);
var rule2 = SpreadsheetApp.newConditionalFormatRule()
.whenTextContains("NA")
    .setBackground("#fce5cd")
    .setRanges([GPTNARNG])
    .build();


var rules = sheet.getConditionalFormatRules();
 rules.push(rule1,rule2);
sheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getActiveSpreadsheet().toast("Autonomous formula have been updated", "👍 Formula Updated", 5);
colorHeaders_NA();
  }
  else {
SpreadsheetApp.getActiveSpreadsheet().toast("Invalid selection. Please choose 'Basic Formulas' or 'Advanced Formulas'.", "Error", 5);
    return;
  }
}

catch (e) {
  handleError(e);  
 logLibraryUsage('Update Formula Common', 'Fail', e.toString()); 
}

}

function unhideAllColumns_NA(sheet) {
  var lastColumn = sheet.getLastColumn();
  sheet.unhideColumn(sheet.getRange(1, 1, 1, lastColumn));
  Logger.log("All columns have been unhidden.");
}

function hideColumns_NA(sheet, headers, columnsToHide) {
 try {
  Logger.log("Sheet: " + sheet.getName());
  Logger.log("Headers: " + headers);
  Logger.log("Columns to Hide: " + columnsToHide);

  if (!headers || headers.length === 0) {
    Logger.log("Headers are undefined or empty.");
    return;
  }

  // Loop through the headers and hide the matching columns
  columnsToHide.forEach(columnName => {
    const columnIndex = headers.indexOf(columnName) + 1;
    if (columnIndex > 0) {
      sheet.hideColumns_NA(columnIndex);
      Logger.log("Hiding column: " + columnName + " at index: " + columnIndex);
    } else {
      Logger.log("Column not found: " + columnName);
    }
  });}
catch (e) {
  handleError(e);
  
}}


function getColumnName_NA(columnIndex) {
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


function colorHeaders_NA() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Assuming headers are in the first row

  // Define headers and their corresponding colors
  const blueHeaders = [	
    "Policy Testing Status",
	"Test Cases Count",
	"Pass",
	"Fail",
	"Need to be Tested",
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
	
  ];
  const redHeaders = [
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
  ];
  const blueColor = "#c9daf8";
  const redColor = "#F4CCCC";

  // Iterate through the headers and apply colors
  headers.forEach((header, index) => {
    const column = index + 1; // Column index is 1-based
    if (blueHeaders.includes(header)) {
      sheet.getRange(1, column).setBackground(blueColor);
    } else if (redHeaders.includes(header)) {
      sheet.getRange(1, column).setBackground(redColor);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("Headers have been colored successfully!");
}

