function compareWithUAT() {
  compareCategoryMaster('UAT_Category_master', 'Seprated Data');
}

function compareWithPROD() {
  compareCategoryMaster('PROD_Category_master','Seprated Data');
}

function compareWithPRODvsUAT() {
  compareCategoryMaster('PROD_Category_master','UAT_Category_master');
}



function compareCategoryMaster(SourceSheetName, targetSheetName) {
 try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
showProgressToast(ss, "Started processing for: " + SourceSheetName);

  const targetSheet = ss.getSheetByName(targetSheetName);
  const categoryMasterSheet = ss.getSheetByName(SourceSheetName);
  const resultSheet = ss.getSheetByName(SourceSheetName + 'Vs.' + targetSheetName +'_verification') || ss.insertSheet(SourceSheetName + 'Vs.' + targetSheetName +'_verification');
  
  // Clear existing data in result sheet except header
  resultSheet.clear();
  ss.setActiveSheet(resultSheet);
  // Add headers to result sheet
  const headers = [
    'Row Number (' + targetSheetName + ')',
    'Row Number (' + SourceSheetName + ')',
    'Department of ' + targetSheetName,
    'Department of ' + SourceSheetName,
    'Category',
    'Sub Categories',
    'Status',
    'Category Created by',
    'Sub Category Created by'
  ];
  resultSheet.appendRow(headers);
  
  // Get headers and data from both sheets
  const sepratedData = targetSheet.getDataRange().getValues();
  const categoryMasterData = categoryMasterSheet.getDataRange().getValues();

  // Find column indexes dynamically
  const sepratedHeaders = sepratedData[0];
  const masterHeaders = categoryMasterData[0];

  const sepratedDeptIdx = sepratedHeaders.indexOf('Department');
  const sepratedCategoryIdx = sepratedHeaders.indexOf('Category');
  const sepratedSubCategoryIdx = sepratedHeaders.indexOf('Sub Categories');

  const masterDeptIdx = masterHeaders.indexOf('Department');
  const masterCategoryIdx = masterHeaders.indexOf('Category');
  const masterSubCategoryIdx = masterHeaders.indexOf('Sub Categories');
  const masterCreatedByIdx = masterHeaders.indexOf('categoryCreatedBy');
  const masterSubCategoryCreatedByIdx = masterHeaders.indexOf('subCategoryCreatedBy');

// || masterCreatedByIdx === -1 removed
  if (sepratedDeptIdx === -1 || sepratedCategoryIdx === -1 || sepratedSubCategoryIdx === -1 ||
      masterDeptIdx === -1 || masterCategoryIdx === -1 || masterSubCategoryIdx === -1 ) {
    throw new Error('One or more required columns are missing.');
  }

  const resultData = [];
  const matchedMasterRows = new Set();
  const BATCH_SIZE = 1000; // Define your batch size

  // Compare rows in batches
  for (let i = 1; i < sepratedData.length; i += BATCH_SIZE) {
    Logger.log("Processing Batch: " + Math.ceil(i / BATCH_SIZE)); // Log batch number
    const end = Math.min(i + BATCH_SIZE, sepratedData.length);
    
    for (let k = i; k < end; k++) {
      const sepRow = sepratedData[k];
      let matchFound = false;

      for (let j = 1; j < categoryMasterData.length; j++) {
        const masterRow = categoryMasterData[j];

        if (sepRow[sepratedDeptIdx] === masterRow[masterDeptIdx] &&
            sepRow[sepratedCategoryIdx] === masterRow[masterCategoryIdx] &&
            sepRow[sepratedSubCategoryIdx] === masterRow[masterSubCategoryIdx]) {
          matchFound = true;
          matchedMasterRows.add(j);
          resultData.push([
            k + 1,
            j + 1,
            sepRow[sepratedDeptIdx],
            masterRow[masterDeptIdx],
            sepRow[sepratedCategoryIdx],
            sepRow[sepratedSubCategoryIdx],
            'Match',
            masterRow[masterCreatedByIdx] || '', // Handle possible missing data
            masterRow[masterSubCategoryCreatedByIdx] || '' // Handle possible missing data
          ]);
        }
      }

      if (!matchFound) {
        resultData.push([
          k + 1,
          '',
          sepRow[sepratedDeptIdx],
          '',
          sepRow[sepratedCategoryIdx],
          sepRow[sepratedSubCategoryIdx],
          'Not available in ' + SourceSheetName,
          '',
          ''
        ]);
      }
    }

    // Write results for this batch to the result sheet
    if (resultData.length > 0) {
      resultSheet.getRange(resultSheet.getLastRow() + 1, 1, resultData.length, headers.length).setValues(resultData);
      resultData.length = 0; // Clear the resultData for the next batch
    }
  }

  // Check for rows in master sheet that are not in Seprated Data
  for (let j = 1; j < categoryMasterData.length; j++) {
    if (!matchedMasterRows.has(j)) {
      const masterRow = categoryMasterData[j];
      resultData.push([
        '',
        j + 1,
        '',
        masterRow[masterDeptIdx],
        masterRow[masterCategoryIdx],
        masterRow[masterSubCategoryIdx],
        'Not available in ' + targetSheetName,
        masterRow[masterCreatedByIdx] || '', // Handle possible missing data
        masterRow[masterSubCategoryCreatedByIdx] || '' // Handle possible missing data
      ]);
    }
  }

  // Write remaining results for master sheet
  if (resultData.length > 0) {
    resultSheet.getRange(resultSheet.getLastRow() + 1, 1, resultData.length, headers.length).setValues(resultData);
  }
 
 showProgressToast(ss, "Completed processing for: " + SourceSheetName);
  } catch (error) {
    Logger.log("Error during comparison: " + error.message);
    ss.toast("Error: " + error.message, 'Error', 10);
  }
}


// function showProgressToast(ss, message) {
//   try {
//     ss.toast(message, 'Progress', 5); // Display for 5 seconds
//     SpreadsheetApp.flush(); // Ensure the UI updates are pushed out immediately
//   } catch (error) {
//     Logger.log("Error in showProgressToast: " + error.message);
//   }
// }
