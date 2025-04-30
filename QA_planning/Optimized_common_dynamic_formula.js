function workLMDynamicFormula() {
  try {
    setFormulaCommon("WorkLM", getWorkLMHiddenColumns(), getWorkLMFormulas);
  } catch (e) {
    handleError(e);
  }
}

function autonomousDynamicFormula() {
  try {
    setFormulaCommon("Autonomous", getAutonomousHiddenColumns(), getAutonomousFormulas);
  } catch (e) {
    handleError(e);
  }
}

function setFormulaCommon(selection, hiddenColumns, formulasGenerator) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QA_Report');
    if (!sheet) throw new Error("Sheet 'QA_Report' not found.");
    
    unhideAllColumns(sheet);

    // Retrieve headers dynamically
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers || headers.length === 0) throw new Error("Headers not found in the sheet.");

    var headerToSearch = "Policy Name";
    var columnIndex = headers.indexOf(headerToSearch) + 1;

    if (columnIndex === 0) {
      Logger.log(headerToSearch + " column not found.");
      return;
    }

        // Apply percentage formatting for specific columns
    headers.forEach((header, index) => {
     // if (header.includes('%')) {
      if (["% Response from GPT", "GPT Pass %", "GPT Fail %", "Fallback %"].includes(header)) {
        let columnIndex = index + 1; // Convert to 1-based index
        sheet.getRange(2, columnIndex, sheet.getMaxRows() - 1).setNumberFormat("0.00%"); // Format as percentage
      }
    });

    // Hide specific columns based on selection
    hideColumns(sheet, headers, hiddenColumns);

    // Clear formats and set formulas
//    clearFormats(sheet, headers, ["Need to be Tested", "% Response from GPT"]);
    const formulas = formulasGenerator(headers); // Pass headers to the formulas generator
    applyFormulas(sheet, headers, formulas);

    // Apply conditional formatting
    applyConditionalFormatting(sheet, headers);

    SpreadsheetApp.getActiveSpreadsheet().toast(`${selection} formulas have been updated`, "👍 Formula Updated", 5);
    // colorHeaders();
  } catch (e) {
    handleError(e);
  }
}

function getColumnName(index) {
  if (index < 1) throw new Error("Invalid column index: " + index);
  let colName = "";
  while (index > 0) {
    let remainder = (index - 1) % 26;
    colName = String.fromCharCode(65 + remainder) + colName;
    index = Math.floor((index - 1) / 26);
  }
  return colName;
}

// When select as autonomous
function getWorkLMHiddenColumns() {
  try {
  return [
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
}catch (e) {
    handleError(e);
  }
  }

//When select as WorkLM
function getAutonomousHiddenColumns() {
  
  try {
    return [
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
}
catch (e) {
    handleError(e);
  }
  }

function getWorkLMFormulas(headers) {
  try {
  validateHeaders(headers, [
    "Policy Name",
    "Test Cases Count",
    "Pass",
    "Fail",
    "Need to be Tested",
    "% Response from GPT"
  ]);

  return {
      // Policy Status
    "Policy Testing Status": `=IFS(${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=0,"WIP",${getColumnName(headers.indexOf("Need to be Tested") + 1)}2>0,"WIP",${getColumnName(headers.indexOf("% Response from GPT") + 1)}2="NA","WIP",${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=${getColumnName(headers.indexOf("Pass") + 1)}2+${getColumnName(headers.indexOf("Fail") + 1)}2,"Done")`,

// Test Cases count
  
  //  "Test Cases Count":`=IF(${getColumnName(headers.indexOf("Policy Name") + 1)}2="",0,COUNTA(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!B2:B")))`,

// below updated formula for Test case count

    "Test Cases Count": `=IF(OR(${getColumnName(headers.indexOf("Policy Name") + 1)}2="", ISERROR(INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!A1"))), 0, COUNTA(INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!" & ADDRESS(2, MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)) & ":" & ADDRESS(ROWS(INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!A:A")), MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)))))`,

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
}
catch (e) {
    handleError(e);
  }
}

function getAutonomousFormulas(headers) {
  try {
  validateHeaders(headers, [
    "Policy Name",
    "Test Cases Count",
    "Pass",
    "Fail",
    "Need to be Tested",
    "Bot stuck",
    "Incorrect audience"
  ]);

  return {
    // Policy Status
    "Policy Testing Status": `=IFS(${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=0,"WIP",${getColumnName(headers.indexOf("Need to be Tested") + 1)}2>0,"WIP",${getColumnName(headers.indexOf("% Response from GPT") + 1)}2="NA","WIP",${getColumnName(headers.indexOf("Test Cases Count") + 1)}2=${getColumnName(headers.indexOf("Pass") + 1)}2+${getColumnName(headers.indexOf("Fail") + 1)}2,"Done")`,

// Test Cases count
  //  "Test Cases Count":`=IF(${getColumnName(headers.indexOf("Policy Name") + 1)}2="",0,COUNTA(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!B2:B")))`,
// below formula updated
"Test Cases Count": `=IF(OR(${getColumnName(headers.indexOf("Policy Name") + 1)}2="", ISERROR(INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!A1"))), 0, COUNTA(INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!" & ADDRESS(2, MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)) & ":" & ADDRESS(ROWS(INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!A:A")), MATCH("Test Case/Query", INDIRECT("'" & ${getColumnName(headers.indexOf("Policy Name") + 1)}2 & "'!1:1"), 0)))))`,
  // Pass
    "Pass": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName(headers.indexOf("Pass") + 1)}$1)`,
    // Fail
    "Fail": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!D:D"),$${getColumnName(headers.indexOf("Fail") + 1)}$1)`,
    // Need to be Tested
    "Need to be Tested": `=${getColumnName(headers.indexOf("Test Cases Count") + 1)}2-${getColumnName(headers.indexOf("Pass") + 1)}2-${getColumnName(headers.indexOf("Fail") + 1)}2`,
    "Bot stuck": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!E:E"),$${getColumnName(headers.indexOf("Bot stuck") + 1)}$1)`,
    "Incorrect audience": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("Incorrect audience") + 1)}$1)`,

    "Incorrect response": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("Incorrect response") + 1)}$1)`,
  
    "hallucination issue": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("hallucination issue") + 1)}$1)`,

    "Policy search": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("Policy search") + 1)}$1)`,
    
    "No/Fallback response": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("No/Fallback response") + 1)}$1)`,
    
    "Partial response given": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("Partial response given") + 1)}$1)`,
    
    "Generic response": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("Generic response") + 1)}$1)`,
    
    "Incorrect language": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("Incorrect language") + 1)}$1)`,
    
    "Incorrect Highlighting": `=COUNTIF(INDIRECT("'"&${getColumnName(headers.indexOf("Policy Name") + 1)}2&"'!F:F"),$${getColumnName(headers.indexOf("Incorrect Highlighting") + 1)}$1)`
};
}catch (e) {
    handleError(e);
  }
}
function validateHeaders(headers, requiredHeaders) {
  requiredHeaders.forEach(header => {
    if (!headers.includes(header)) throw new Error(`Missing required header: ${header}`);
  });
}

function clearFormats(sheet, headers, columns) {
  columns.forEach(column => {
    const range = sheet.getRange(`${getColumnName(headers.indexOf(column) + 1)}:${getColumnName(headers.indexOf(column) + 1)}`);
    range.clearFormat();
    Logger.log("Cleared format for " + column);
  });
}

function applyFormulas(sheet, headers, formulas) {
try {
    for (var header in formulas) {
    const columnIndex = headers.indexOf(header) + 1;
    if (columnIndex > 0) {
      const formula = formulas[header];
      const cellBelowHeader = sheet.getRange(2, columnIndex);
      cellBelowHeader.setFormula(formula).setBorder(true, true, true, true, true, true, "SOLID", null);
    }
  }
}
catch (e) {
    handleError(e);
  }
}

function applyConditionalFormatting(sheet, headers) {
  try {
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#f4cccc")
      .setRanges([sheet.getRange(`${getColumnName(headers.indexOf("Need to be Tested") + 1)}:${getColumnName(headers.indexOf("Need to be Tested") + 1)}`)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("NA")
      .setBackground("#fce5cd")
      .setRanges([sheet.getRange(`${getColumnName(headers.indexOf("% Response from GPT") + 1)}:${getColumnName(headers.indexOf("% Response from GPT") + 1)}`)])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
}
catch (e) {
    handleError(e);
  }
}



function unhideAllColumns(sheet) {
  var lastColumn = sheet.getLastColumn();
  sheet.unhideColumn(sheet.getRange(1, 1, 1, lastColumn));
  Logger.log("All columns have been unhidden.");
}

function hideColumns(sheet, headers, columnsToHide) {
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
      sheet.hideColumns(columnIndex);
      Logger.log("Hiding column: " + columnName + " at index: " + columnIndex);
    } else {
      Logger.log("Column not found: " + columnName);
    }
  });}
catch (e) {
  handleError(e);
  
}}

