function setFormulaByDynamicHeader() {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QA_Report'); // Replace 'YourSheetName' with the name of your sheet.

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
  var formulas = {
    // Policy Status
    "Policy Testing Status": `=IFS(${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=0,"WIP",${getColumnName(headers.indexOf("Need to be Tested") + 1)}2>0,"WIP",${getColumnName(headers.indexOf("% Response from GPT") + 1)}2="NA","WIP",${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=${getColumnName(headers.indexOf("Pass") + 1)}2+${getColumnName(headers.indexOf("Fail") + 1)}2,"Done")`,
// Test Cases count
    "Test Cases Count":`=IF(${getColumnName(headers.indexOf("Policy Name") + 1)}2="",0,COUNTA(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!B2:B")))`,
    // Pass
    "Pass": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName(headers.indexOf("Pass") + 1)}$1)`,
    // Fail
    "Fail": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName(headers.indexOf("Fail") + 1)}$1)`,
    // Need to be Tested
    "Need to be Tested": `=${getColumnName(headers.indexOf("Test Cases Count") + 1)}2-${getColumnName(headers.indexOf("Pass") + 1)}2-${getColumnName(headers.indexOf("Fail") + 1)}2`,
   // Response Count from GPT
    "Response Count from GPT" : `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"GPT")`,
    // % Response from GPT
    "% Response from GPT" : `=if(${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=0,"NA",(${getColumnName(headers.indexOf("Response Count from GPT") + 1)}2/${getColumnName(headers.indexOf("Test Cases Count") + 1)}2))`, // removed *100)    
    // GPT Pass %    
    "GPT Pass %" : `=if(${getColumnName(headers.indexOf("GPT Pass Count") + 1)}2=0,"NA",(${getColumnName(headers.indexOf("GPT Pass Count") + 1)}2/${getColumnName(headers.indexOf("Response Count from GPT") + 1)}2))`, // removed *100)
    
    
    // GPT Fail %
    "GPT Fail %" : `=if(${getColumnName(headers.indexOf("GPT Fail Count") + 1)}2=0,"0",(${getColumnName(headers.indexOf("GPT Fail Count") + 1)}2/${getColumnName(headers.indexOf("Response Count from GPT") + 1)}2))`, // removed *100)
    // Fallback Count
    "Fallback Count" : `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"Fallback")`,
    // FallBack
    "Fallback %" : `=if(${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=0,"NA",(${getColumnName(headers.indexOf("Fallback Count") + 1)}2/${getColumnName(headers.indexOf("Test Cases Count") + 1)}2))`, // removed *100
    // response from search count
    "response from search count" : `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"Search")`,
    // % response from Search
    "% response from Search" : `=if(${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=0,"NA",(${getColumnName(headers.indexOf("response from search count") + 1)}2/${getColumnName(headers.indexOf("Test Cases Count") + 1)}2))`, // removed *100
    // GPT Pass Count
    "GPT Pass Count" : `=COUNTIFs(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName(headers.indexOf("Pass") + 1)}$1,INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"GPT")`,
    // GPT Fail Count
    "GPT Fail Count" : `=COUNTIFs(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName(headers.indexOf("Fail") + 1)}$1,INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),"GPT")`
  
  };

sheet.getRange(`${getColumnName(headers.indexOf("Need to be Tested")+1)}:${getColumnName(headers.indexOf("Need to be Tested")+1)}`).clearFormat();
Logger.log("Clear Condition format "  + `${getColumnName(headers.indexOf("Need to be Tested")+1)}:${getColumnName(headers.indexOf("Need to be Tested")+1)}`);
sheet.getRange(`${getColumnName(headers.indexOf("% Response from GPT")+1)}:${getColumnName(headers.indexOf("% Response from GPT")+1)}`).clearFormat();
Logger.log("Clear Condition format " + `${getColumnName(headers.indexOf("% Response from GPT")+1 )}:${getColumnName(headers.indexOf("% Response from GPT")+1)}`);


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


var nottestedRNG = sheet.getRange(`${getColumnName(headers.indexOf("Need to be Tested")+1)}:${getColumnName(headers.indexOf("Need to be Tested")+1)}`);
var rule1 = SpreadsheetApp.newConditionalFormatRule()
.whenNumberGreaterThan(0)
    .setBackground("#f4cccc")
    .setRanges([nottestedRNG])
    .build();
    
// GPT Range condition Format
var GPTNARNG = sheet.getRange(`${getColumnName(headers.indexOf("% Response from GPT")+1)}:${getColumnName(headers.indexOf("% Response from GPT")+1)}`);
var rule2 = SpreadsheetApp.newConditionalFormatRule()
.whenTextContains("NA")
    .setBackground("#fce5cd")
    .setRanges([GPTNARNG])
    .build();


var rules = sheet.getConditionalFormatRules();
 rules.push(rule1,rule2);
sheet.setConditionalFormatRules(rules);

SpreadsheetApp.getActiveSpreadsheet().toast( "Kindly verify the same!" , "👍 Formula Updated", 5 )

}

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

