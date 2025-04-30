function separateRecordsForPROD() {
  separateRecordsBySubCategories('PROD');
}

function separateRecordsForUAT() {
  separateRecordsBySubCategories('UAT');
}

function separateRecordsForREQ() {
  separateRecordsBySubCategories('REQ');
}


function separateRecordsBySubCategories(env) {
try{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName, targetSheetName;

  if (env === 'PROD') {
    sourceSheetName = 'PROD Assignee Config';
    targetSheetName = 'PROD Assignee Config separated';
  } else if (env === 'UAT') {
    sourceSheetName = 'UAT Assignee Config';
    targetSheetName = 'UAT Assignee Config separated';
  } else if (env === 'REQ') {
    sourceSheetName = 'Requirement Assignee Config';
    targetSheetName = 'Requirement Assignee Config separated';
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('Invalid environment specified', '⚠️ Warning', 10);
    return;
  }

  var sourceSheet = ss.getSheetByName(sourceSheetName);

  // Delete the target sheet if it already exists
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    ss.deleteSheet(targetSheet);
  }

  // Create a new sheet for separated records
  targetSheet = ss.insertSheet(targetSheetName);  // Reassigning targetSheet to new sheet

  // Get header row to find the column index of "Sub Categories"
  var headerRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  var subCategoriesColumnIndex = headerRow.indexOf('Sub Categories') + 1;

  // Check if "Sub Categories" header is found
  if (subCategoriesColumnIndex === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Header Sub Categories not found', '⚠️ Warning', 10);
    return;
  }

  // Get data from the source sheet
  var data = sourceSheet.getDataRange().getValues();

  // Prepare target data array
  var targetData = [];

  // Loop through the data and separate records based on Sub Categories
  data.forEach(function(row) {
    var subCategories = row[subCategoriesColumnIndex - 1].split(';');
    subCategories.forEach(function(subCategory) {
      var newRow = row.slice(); // Copy the existing row
      newRow[subCategoriesColumnIndex - 1] = subCategory.trim(); // Replace Sub Categories with the individual value
      targetData.push(newRow); // Add the new row to the target data array
    });
  });

  // Append the target data array to the target sheet
  if (targetData.length > 0) {
    targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);
    Logger.log('Records separated successfully.');
  } else {
    Logger.log('No records found to separate.');
  }
}
  catch(error)
  {
    handleError(error);
  }
}
