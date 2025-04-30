function BKP_importXLSXFromLinksAndValidateAndUpdateStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketingFormSheet = ss.getSheetByName('TicketingForm');
  if (!ticketingFormSheet) {
    Logger.log("TicketingForm sheet not found");
    return;
  }


  const data = ticketingFormSheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("No data found in TicketingForm sheet");
    return;
  }

  // Find the column indices for Name and Link
  const headers = data[0];
  const nameColumnIndex = headers.indexOf('Name');
  const linkColumnIndex = headers.indexOf('Link');
  let statusColumnIndex = headers.indexOf('Form Validation Status');

  if (nameColumnIndex === -1 || linkColumnIndex === -1) {
    Logger.log("Name or Link column not found. Headers: " + headers.join(", "));
    return;
  }
  Logger.log("Name column index: " + nameColumnIndex);
  Logger.log("Link column index: " + linkColumnIndex);

  // If Form Validation Status column doesn't exist, add it as the last column
  if (statusColumnIndex === -1) {
    statusColumnIndex = headers.length;
    ticketingFormSheet.getRange(1, statusColumnIndex + 1).setValue('Form Validation Status');
  }

  const exportedDataSheet = ss.getSheetByName('Exported Data from ticketing forms') || ss.insertSheet('Exported Data from ticketing forms');
  const exportedHeaders = ['Category', 'Sub Categories', 'Form Name', 'Extra Field 1', 'Extra Field 2','Validation Status'];
   
    // Clear existing data (excluding headers)
  if (exportedDataSheet.getLastRow() > 1) {
    exportedDataSheet.getRange(2, 1, exportedDataSheet.getLastRow() - 1, exportedHeaders.length).clearContent();
    exportedDataSheet.clearConditionalFormatRules();
    
    Logger.log("Exported data sheet clear.")
  }

  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameColumnIndex];
    const url = data[i][linkColumnIndex];

    Logger.log("Processing row " + (i + 1) + ": Name = " + name + ", URL = " + url);
    SpreadsheetApp.getActiveSpreadsheet().toast("Processing row " + (i + 1) + ": Name = " + name,"Progress",5)

    if (name && url) {
      try {
        Logger.log("Fetching URL: " + url);
        
        const response = UrlFetchApp.fetch(url);
        Logger.log("URL fetched successfully");
        const blob = response.getBlob().setContentType(MimeType.MICROSOFT_EXCEL);
        
        // Generate file name based on column values
        const fileName = `${name.trim()}.xlsx`;
        Logger.log("Generated file name: " + fileName);
        
        // Create file in Drive
        const file = DriveApp.createFile(blob);
        file.setName(fileName);
        Logger.log("File created with ID: " + file.getId());

        const fileId = file.getId();
        
        // Use the Advanced Drive Service to convert the file to Google Sheets format
        const tempSpreadsheet = Drive.Files.copy({}, fileId, {
          convert: true
        });
        const tempFileId = tempSpreadsheet.id;
        Logger.log("Temporary spreadsheet created with ID: " + tempFileId);

        // Open the converted Google Sheet
        const tempSheet = SpreadsheetApp.openById(tempFileId).getSheets()[0];
        Logger.log("Temporary sheet opened");
        // Check if a sheet with the same name already exists
        let importSheet = ss.getSheetByName(name);
        if (!importSheet) {
          // If the sheet doesn't exist, create a new one
          importSheet = ss.insertSheet(name);
          Logger.log("Sheet created with name: " + name);
        } else {
          // If the sheet exists, clear its content
          importSheet.clear();
          Logger.log("Sheet cleared with name: " + name);
        }
        
        // Copy data from the temporary sheet to the import sheet
        const tempData = tempSheet.getDataRange().getValues();
        importSheet.getRange(1, 1, tempData.length, tempData[0].length).setValues(tempData);
        Logger.log("Data copied to sheet: " + name);

        // freez first row and column
        importSheet.setFrozenColumns(1);
        importSheet.setFrozenRows(1);

        // Delete the temporary files
        DriveApp.getFileById(tempFileId).setTrashed(true);
        DriveApp.getFileById(fileId).setTrashed(true);
        Logger.log("Temporary files deleted");

        // Call validation function
        const validatedSheetName = `${name.trim()} validated`;
        const { validationStatus, validatedData } = BKP_validateTicketFormBothCM_1(importSheet, validatedSheetName);
        
        // Update Form Validation Status in TicketingForm sheet
        ticketingFormSheet.getRange(i + 1, statusColumnIndex + 1).setValue(validationStatus);

        // Append validated data to the "Exported Data from ticketing forms" sheet
        BKP_appendValidatedData(exportedDataSheet, validatedData, name);
        
      } catch (e) {
        Logger.log(`Error processing ${name} with URL ${url}: ${e.toString()}`);
        // Set status to indicate error in case of exception
        ticketingFormSheet.getRange(i + 1, statusColumnIndex + 1).setValue('Error: ' + e.toString());
      }
    }
  }

  // Add conditional formatting to the exported data sheet
  BKP_addConditionalFormatting(exportedDataSheet);
}

function BKP_validateTicketFormBothCM_1(importSheet, validatedSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear existing validated sheet if it exists
  const existingValidatedSheet = ss.getSheetByName(validatedSheetName);
  if (existingValidatedSheet) {
    ss.deleteSheet(existingValidatedSheet);
  }
  
  // Create a new validated sheet
  const validatedSheet = ss.insertSheet(validatedSheetName);
  validatedSheet.appendRow(['Category', 'Sub Categories','Status']); // Header
  
  const lastRow = importSheet.getLastRow();
  
  // Check if the import sheet is blank
  if (lastRow === 0) {
    validatedSheet.appendRow(['No data found in import sheet', 'Error']);
    return { validationStatus: 'No data found', validatedData: [] };
  }
  
  // Get the header row to dynamically find column indices
  const headerRow = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues()[0];
  const dependsOnColumnIndex = headerRow.indexOf('dependsOnValue') + 1;
  const optionsColumnIndex = headerRow.indexOf('options') + 1;
  const keyColumnIndex = headerRow.indexOf('key') + 1;
  
  // Get values from the import sheet
  const dependsOnValues = importSheet.getRange(2, dependsOnColumnIndex, lastRow - 1, 1).getValues().flat();
  const keyValues = importSheet.getRange(2, keyColumnIndex, lastRow - 1, 1).getValues().flat();
  
  // Search for the row index containing the key "category"
  let categoryRowIndex = keyValues.indexOf("category");
  if (categoryRowIndex === -1) {
    categoryRowIndex = 1; // If "category" is not found from row 3, we set it to 1
  }
  
  // Calculate the actual row number in the spreadsheet where the "category" key is found
  const categoryRow = categoryRowIndex + 2; // Add 2 because we start counting from row 2 in the spreadsheet
  
  // Retrieve options value from the row where "category" key is found
  const options = importSheet.getRange(categoryRow, optionsColumnIndex).getValue();
  const optionsArray = options.split("||").map(option => option.trim());
  Logger.log("Options picked from 'options' column: " + optionsArray.join(", "));
  
  if (keyValues.includes("subCategory")) {
    SpreadsheetApp.getActiveSpreadsheet().toast('New CM form detected: Based on key value "subCategory"', '⚠️ Form Details', 5);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('Mono CM form detected: Based on key value "category.subCategory.name"', '⚠️ Form Details', 5);
  }
  

  // Create a mapping of subcategories to categories (options)
  const subCategoryMap = {};
  for (let i = 0; i < keyValues.length; i++) {
    const key = keyValues[i];
    if (key === "subCategory" || key === "category.subCategory.name") {
      const subCategoryValue = dependsOnValues[i];
      if (subCategoryValue) {
        const subCategoryOptions = importSheet.getRange(i + 2, optionsColumnIndex).getValue();
        subCategoryMap[subCategoryValue] = subCategoryOptions; // Store subcategory options as is
      }
    }
  }

  let validationStatus = 'Pass'; // Default status
  const validatedData = [];
  
  // Check the status for each value where the key matches "category.subCategory.name"
  for (let i = 0; i < dependsOnValues.length; i++) {
    const key = keyValues[i];
    // for new cm SubCategory and mono cm category.SubCategoryName
    if (key === "subCategory" || key === "category.subCategory.name") {
      const value = dependsOnValues[i];
      const status = optionsArray.includes(value) ? 'Available in Both' : 'Missing in Column options; Its critical';
      const subCategory = subCategoryMap[value] || '';
      validatedSheet.appendRow([value, subCategory, status]);
      validatedData.push([value, subCategory, status]);      
      // Update validation status based on the worst-case scenario
      if (status === 'Missing in Column options; Its critical') {
        validationStatus = 'Critical error';
      } else if (status === 'Missing in Column dependsOnValue') {
        validationStatus = 'Need to check';
      }
    }
  }
  
  // Check for values in "options" column that are not in "dependsOnValue" column
  if (optionsArray) {
    optionsArray.forEach(option => {
      if (!dependsOnValues.includes(option)) {
          validatedSheet.appendRow([option, '', 'Missing in Column dependsOnValue']);
          validatedData.push([option, '', 'Missing in Column dependsOnValue']);
         validationStatus = 'Need to check'; // Update status if any missing dependency
      }
    });
  }
  
  return { validationStatus, validatedData };

}

function BKP_appendValidatedData(exportedDataSheet, validatedData, formName) {
  if (validatedData.length === 0) return;

  const headers = exportedDataSheet.getRange(1, 1, 1, exportedDataSheet.getLastColumn()).getValues()[0];
  const categoryIndex = headers.indexOf('Category');
  const subCategoryIndex = headers.indexOf('Sub Categories');
  const formNameIndex = headers.indexOf('Form Name');
  const extraField1Index = headers.indexOf('Extra Field 1');
  const extraField2Index = headers.indexOf('Extra Field 2');
 const validationStatusIndex = headers.indexOf('Validation Status');


  validatedData.forEach(row => {
    const category = row[0];
    const subCategory = row[1];
    const status = row[2];
    const newRow = Array(headers.length).fill('');
    newRow[categoryIndex] = category;
    newRow[subCategoryIndex] = subCategory;
    newRow[formNameIndex] = formName;
    newRow[validationStatusIndex] = status;  // Write the status to the new 'Validation Status' column
    exportedDataSheet.appendRow(newRow);
  });
}

function BKP_addConditionalFormatting(sheet) {
  // Get the headers from the first row
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column indices for "Category", "Sub Categories", and "Validation Status"
  const categoryColumnIndex = headers.indexOf('Category') + 1; // +1 because getRange is 1-based
  const subCategoryColumnIndex = headers.indexOf('Sub Categories') + 1; // +1 because getRange is 1-based
  const validationStatusColumnIndex = headers.indexOf('Validation Status') + 1; // +1 because getRange is 1-based

  if (categoryColumnIndex === 0 || subCategoryColumnIndex === 0 || validationStatusColumnIndex === 0) {
    Logger.log("Category, Sub Categories, or Validation Status column not found");
    return;
  }

  // Clear existing conditional formatting rules
  Logger.log("Clearing Exported sheet formating");
  sheet.clearConditionalFormatRules();

  // Define the ranges for conditional formatting
  const categoryRange = sheet.getRange(2, categoryColumnIndex, sheet.getLastRow() - 1);
  const subCategoryRange = sheet.getRange(2, subCategoryColumnIndex, sheet.getLastRow() - 1);

  // Create conditional formatting rule for "Category" column
  const ruleCategory = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISNUMBER(SEARCH(";", INDIRECT(ADDRESS(ROW(), ${categoryColumnIndex}))))`)
    .setBackground('#ffcccc')  // Light red background
    .setRanges([categoryRange])
    .build();

  // Create conditional formatting rule for "Sub Categories" column
  const ruleSubCategory = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISNUMBER(SEARCH(";", INDIRECT(ADDRESS(ROW(), ${subCategoryColumnIndex}))))`)
    .setBackground('#ffcccc')  // Light red background
    .setRanges([subCategoryRange])
    .build();

  // Apply the new rules to the sheet
  const rules = [ruleCategory, ruleSubCategory];
  sheet.setConditionalFormatRules(rules);
  // set first row as freeze
  sheet.setFrozenRows(1);
  // Update the Validation Status column if semicolon is found
  const lastRow = sheet.getLastRow();
  const categoryValues = sheet.getRange(2, categoryColumnIndex, lastRow - 1).getValues();
  const subCategoryValues = sheet.getRange(2, subCategoryColumnIndex, lastRow - 1).getValues();
  const validationStatusRange = sheet.getRange(2, validationStatusColumnIndex, lastRow - 1);

  const validationStatusValues = validationStatusRange.getValues();

  for (let i = 0; i < categoryValues.length; i++) {
    const categoryValue = categoryValues[i][0];
    const subCategoryValue = subCategoryValues[i][0];
    
    if (categoryValue.includes(";") || subCategoryValue.includes(";")) {
      validationStatusValues[i][0] = 'Category or Sub category has ; Semicolon';
    }
  }

  validationStatusRange.setValues(validationStatusValues);
}

