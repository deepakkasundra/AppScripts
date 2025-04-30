function consolidateData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var qaReportSheet = ss.getSheetByName("QA_Report");
    var testCasesFormatSheet = ss.getSheetByName("Test Cases Format");

    if (!qaReportSheet || !testCasesFormatSheet) {
      Logger.log("QA_Report or Test Cases Format sheet not found.");
   
      return;
    }

    // Check if "Consolidated Data" sheet exists
    var consolidatedSheet = ss.getSheetByName("Consolidated Data");
    if (consolidatedSheet) {
      // Show a popup message asking to clear existing data or stop
      var response = Browser.msgBox("Consolidated Data", 
        "The 'Consolidated Data' sheet already exists. Do you want to clear its existing data?", 
        Browser.Buttons.YES_NO);

      if (response == "yes") {
        // Clear existing data if user chose "Yes"
        consolidatedSheet.clearContents();
      } else {
        // Stop the script if user chose "No"
        Logger.log("Consolidation process stopped by the user.");
        return;
      }
    } else {
      // Create the sheet if it does not exist
      consolidatedSheet = ss.insertSheet("Consolidated Data");
    }

    // Get header row from Test Cases Format for dynamic header mapping
    var testCasesFormatHeaders = testCasesFormatSheet.getDataRange().getValues()[0];
    testCasesFormatHeaders = testCasesFormatHeaders.filter(function(header) {
      return header && header !== "Go to Main Page"; // Filter out blank headers and "Go to Main Page"
    });

    // Consolidated headers include "Sheet Name" + headers from Test Cases Format
    var consolidatedHeaders = ["Sheet Name"].concat(testCasesFormatHeaders);

    // Add headers to the "Consolidated Data" sheet
    consolidatedSheet.appendRow(consolidatedHeaders);

    // Get data range and values from QA_Report sheet
    var dataValues = qaReportSheet.getDataRange().getValues();

    // Find the column containing policy names dynamically in QA_Report
    var policyNameColumn = dataValues[0].indexOf("Policy Name");
    if (policyNameColumn === -1) {
      Logger.log("Policy Name column not found.");
      return;
    }

    // Iterate over each row in QA_Report starting from the second row
    for (var i = 1; i < dataValues.length; i++) {
      var policyName = dataValues[i][policyNameColumn];
      Logger.log(policyName);

      // Access the policy sheet for each policy name
      var policySheet = ss.getSheetByName(policyName);
      if (!policySheet) {
        Logger.log("Policy sheet '" + policyName + "' not found.");
 SpreadsheetApp.getActiveSpreadsheet().toast("Policy sheet '" + policyName + "' not found.", "Warning", 5);
        
        // Add a row indicating the policy sheet was not found
        var rowDataNotFound = [policyName, "<sheet Name Not found>"];
        consolidatedSheet.appendRow(rowDataNotFound);
        continue;
      }

      // Get data range and values from the policy sheet
      var policyDataValues = policySheet.getDataRange().getValues();
      var policyHeaders = policyDataValues[0];

      // Map column positions dynamically based on Test Cases Format headers
      var columnMap = {};
      testCasesFormatHeaders.forEach(function(header) {
        var headerIndex = policyHeaders.indexOf(header);
        columnMap[header] = headerIndex !== -1 ? headerIndex : "Header not found";
      });

      // Iterate over each row in the policy sheet starting from the second row
      for (var j = 1; j < policyDataValues.length; j++) {
        var rowData = [policyName]; // Start each row with the policy name

        // Add data based on column map, handling missing headers
        testCasesFormatHeaders.forEach(function(header) {
          var colIndex = columnMap[header];
          rowData.push(colIndex !== "Header not found" ? policyDataValues[j][colIndex] : "Header not found");
        });

        // Append the row to the "Consolidated Data" sheet
        consolidatedSheet.appendRow(rowData);
      }
    }

    // Clear existing conditional formatting rules
    consolidatedSheet.clearConditionalFormatRules();

    // Apply new conditional formatting to highlight "Header not found" cells in light red
    var lastRow = consolidatedSheet.getLastRow();
    var lastCol = consolidatedSheet.getLastColumn();
    var range = consolidatedSheet.getRange(2, 1, lastRow - 1, lastCol); // Exclude header row

    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Header not found")
      .setBackground("#f4cccc") // Light red color
      .setRanges([range])
      .build();

    var rules = consolidatedSheet.getConditionalFormatRules();
    rules.push(rule);
    consolidatedSheet.setConditionalFormatRules(rules);

    Logger.log("Data consolidation and formatting completed successfully.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Data consolidation and formatting completed successfully.", "Success", 5);

  } catch (e) {
    handleError(e);
  logLibraryUsage('Old Consolidated Data', 'Fail', e.toString());
  }
}


