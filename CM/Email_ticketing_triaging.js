function Triaging_Email_ticketing_UAT() {
  Triaging_Email_ticketing('UAT');
}

function Triaging_Email_ticketing_PROD() {
  Triaging_Email_ticketing('PROD');
}
function Triaging_Email_ticketing(environment) {
  var maxResponses = 1; // Maximum number of responses (configurable)

  try {  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var triagingSheet = ss.getSheetByName('Triaging');
    if (!triagingSheet) {
      Logger.log('Triaging sheet not found');
      SpreadsheetApp.getActiveSpreadsheet().toast('Triaging Sheet not available', '⚠️ Further execution stopped.', 10);
      return;
    }

    // Clear content from the "Response category 1" column and subsequent columns
    var triagingHeadersRange = triagingSheet.getRange(1, 1, 1, triagingSheet.getLastColumn());
    var triagingHeadersValues = triagingHeadersRange.getValues()[0];
    var responseCategory1Index = triagingHeadersValues.indexOf('Response category 1');

    var categoryMatchingStatusIndex = triagingHeadersValues.indexOf('Triaging Matching Status');

    if (responseCategory1Index === -1) {
      Logger.log('Header "Response category 1" not found in Triaging sheet');
      SpreadsheetApp.getActiveSpreadsheet().toast('Header "Response category 1" not found in Triaging sheet', '⚠️ Further execution stopped.', 10);
      return;
    }

    // Check if "Triaging Matching Status" column exists
    if (categoryMatchingStatusIndex === -1) {
      // If not found, add a new column for "Triaging Matching Status" at the end
      triagingSheet.insertColumnAfter(triagingSheet.getLastColumn());
      triagingSheet.getRange(1, triagingSheet.getLastColumn()).setValue('Triaging Matching Status');
      categoryMatchingStatusIndex = triagingSheet.getLastColumn() - 1; // Set the new column index
    } else {
      // If the column exists, clear the existing data in "Triaging Matching Status"
      var categoryMatchingStatusRange = triagingSheet.getRange(2, categoryMatchingStatusIndex + 1, triagingSheet.getLastRow() - 1, 1);
      categoryMatchingStatusRange.clearContent();
    }


 // Clear content from the "Response category 1" column and subsequent columns
triagingSheet.getRange(2, responseCategory1Index, triagingSheet.getLastRow() - 1, triagingSheet.getLastColumn() - responseCategory1Index + 1).clearContent();


    // Clear previous conditional formatting
    triagingSheet.getRange(2, categoryMatchingStatusIndex + 1, triagingSheet.getLastRow() - 1, 1).clearFormat();

    // Apply conditional formatting rules
    var rules = triagingSheet.getConditionalFormatRules();

    // Add formatting for "Full Match"
    var fullMatchRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Full Match')
      .setBackground('#C8E6C9') // Green
      .setRanges([triagingSheet.getRange(2, categoryMatchingStatusIndex + 1, triagingSheet.getLastRow() - 1, 1)])
      .build();

    // Add formatting for "Partial : Category Match"
    var categoryMatchRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Partial : Category Match')
      .setBackground('#ffffe0') // Yellow
      .setRanges([triagingSheet.getRange(2, categoryMatchingStatusIndex + 1, triagingSheet.getLastRow() - 1, 1)])
      .build();

    // Add formatting for "Partial : Sub Category Match"
    var subCategoryMatchRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Partial : Sub Category Match')
      .setBackground('#ffcc99') // Light Blue
      .setRanges([triagingSheet.getRange(2, categoryMatchingStatusIndex + 1, triagingSheet.getLastRow() - 1, 1)])
      .build();

    // Add formatting for "Fail"
    var failRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Fail')
      .setBackground('#ffcccb') // Red
      .setRanges([triagingSheet.getRange(2, categoryMatchingStatusIndex + 1, triagingSheet.getLastRow() - 1, 1)])
      .build();

    // Apply all rules to the sheet
    rules.push(fullMatchRule, categoryMatchRule, subCategoryMatchRule, failRule);
    triagingSheet.setConditionalFormatRules(rules);


    var mainSheet = ss.getSheetByName('Main');
    if (!mainSheet) {
      Logger.log('Main sheet not found');
      SpreadsheetApp.getActiveSpreadsheet().toast('Main Sheet not available', '⚠️ Further execution stopped.', 10);
      return;
    }

    // Get header values and find the indexes of required columns in Main sheet
    var mainHeadersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
    var mainHeadersValues = mainHeadersRange.getValues()[0];
    var rowIndex = 2; // as value available at 2

    // Set the BOT ID index based on the environment (UAT or PROD)
    var BOTIDIndex;
    if (environment === 'UAT') {
      BOTIDIndex = mainHeadersValues.indexOf('UAT BOT ID') + 1;
    } else if (environment === 'PROD') {
      BOTIDIndex = mainHeadersValues.indexOf('PROD BOT ID') + 1;
    }

    var NLP_TokenIndex = mainHeadersValues.indexOf('NLP Token') + 1;
    var NLP_URLIndex = mainHeadersValues.indexOf('NLP Dashboard') + 1;

    if (BOTIDIndex === -1 || NLP_TokenIndex === -1 || NLP_URLIndex === -1) {
      Logger.log('Required columns not found in Main sheet');
      SpreadsheetApp.getActiveSpreadsheet().toast('Required columns not found in Main sheet', '⚠️ Further execution stopped.', 10);
      return;
    }

    var BOTID = mainSheet.getRange(rowIndex, BOTIDIndex).getValue();
    var uatJwt_value = mainSheet.getRange(rowIndex, NLP_TokenIndex).getValue();
    var uatJwt = 'Bearer ' + uatJwt_value;
    var NLP_URL = mainSheet.getRange(rowIndex, NLP_URLIndex).getValue();

    if (!BOTID) {
      Logger.log(`Generate ${environment} BOT ID first`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Generate ${environment} BOT ID first`, '⚠️ Further execution stopped.', 10);
      return;
    }

    var triagingHeadersRange = triagingSheet.getRange(1, 1, 1, triagingSheet.getLastColumn());
    var triagingHeadersValues = triagingHeadersRange.getValues()[0];
    var subjectqueryIndex = triagingHeadersValues.indexOf('Subject Line');
    var emailBodyIndex = triagingHeadersValues.indexOf('Email Body');
    var categoryIndex = triagingHeadersValues.indexOf('Category');
    var subCategoryIndex = triagingHeadersValues.indexOf('Sub Category');


    if (subjectqueryIndex === -1 || emailBodyIndex === -1 || categoryIndex === -1 || subCategoryIndex === -1) {
      Logger.log('Required columns not found in Triaging sheet');
      SpreadsheetApp.getActiveSpreadsheet().toast('Required columns not found in Triaging sheet', '⚠️ Further execution stopped.', 10);
      return;
    }

    // Fetching data from Triaging sheet
    var dataRange = triagingSheet.getDataRange();
    var dataValues = dataRange.getValues();

    for (var i = 1; i < dataValues.length; i++) { // Start from 1 to skip header row
      var query = dataValues[i][subjectqueryIndex];
      var emailBody = dataValues[i][emailBodyIndex];

      // Check if both query and email body are empty
      if (!query && !emailBody) {
        Logger.log('Both query and email body are empty in row ' + (i + 1));
        continue;
      }

      var url = `${NLP_URL}/@@@@@@@@/${BOTID}/@@@@@@@@/`;
      var headers = {
        'Accept': '*/*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Authorization': uatJwt,
        'Connection': 'keep-alive',
        'Content-Type': 'application/json'
      };

      var payload = {
        'query': query,
        'categoryList': '',
        'isEmail': true,
        'emailBody': emailBody
      };

      var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload)
      };

      try {
        var response = UrlFetchApp.fetch(url, options);
        var responseData = JSON.parse(response.getContentText());

        Logger.log(responseData); // Log fetched data

        if (responseData && responseData.data && responseData.data.merged_cat && responseData.data.merged_cat.length > 0) {
          var mergedCategories = responseData.data.merged_cat;
          var categoryData = [];
          for (var j = 0; j < Math.min(mergedCategories.length, maxResponses); j++) {
            var matchedCategory = mergedCategories[j][0][0];
            var matchedSubCategory = mergedCategories[j][0][1];
            categoryData.push(matchedCategory, matchedSubCategory);
          }
          // Write data to the sheet for the respective row
          var categoryDataRange = triagingSheet.getRange(i + 1, triagingHeadersValues.indexOf('Response category 1') + 1, 1, Math.min(maxResponses * 2, 4));  // 4 = select 4 columns (Response category 1, Response Sub category 1, Response category 2, and Response Sub category 2).
          categoryDataRange.clearContent();
          categoryDataRange.setValues([categoryData]);

          // Matching logic for Category/Sub-category
          var sheetCategory = dataValues[i][categoryIndex];
          var sheetSubCategory = dataValues[i][subCategoryIndex];
          var status;

          if (sheetCategory === matchedCategory && sheetSubCategory === matchedSubCategory) {
            status = 'Full Match';
          } else if (sheetCategory === matchedCategory) {
            status = 'Partial : Category Match';
          } else if (sheetSubCategory === matchedSubCategory) {
            status = 'Partial : Sub Category Match';
          } else {
            status = 'Fail';
          }

          // Update Category Matching Status
          triagingSheet.getRange(i + 1, categoryMatchingStatusIndex + 1).setValue(status);

          SpreadsheetApp.flush(); // Flush changes to ensure they're written immediately
        }
      } catch (error) {
        Logger.log('Error fetching data: ' + error);
        var errorCell = triagingSheet.getRange(i + 1, triagingHeadersValues.indexOf('Response category 1') + 1);
        errorCell.clearContent(); // Clear data from the cell
        errorCell.setValue('Error: ' + error); // Write error message to the cell

        triagingSheet.getRange(i + 1, categoryMatchingStatusIndex + 1).setValue("Fail " + error);
        SpreadsheetApp.flush(); // Flush changes to ensure they're written immediately
      }
    }
  }
  catch (error) {
    Logger.log('Unexpected error: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Unexpected error: ${error}`, '⚠️ Execution stopped.', 10);
  }
}
