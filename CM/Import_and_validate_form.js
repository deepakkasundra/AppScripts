function importXLSXFromLinksAndValidateAndUpdateStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketingFormSheet = ss.getSheetByName('TicketingForm');
  if (!ticketingFormSheet) {
    Logger.log("TicketingForm sheet not found");
           Browser.msgBox("TicketingForm sheet not found.", Browser.Buttons.OK);

    return;
  }


  const data = ticketingFormSheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("No data found in TicketingForm sheet");
    Browser.msgBox("No data found in TicketingForm sheet.", Browser.Buttons.OK);

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

// Check if the Link column has any data
let isLinkColumnEmpty = true;
for (let i = 1; i < data.length; i++) { // Start from 1 to skip headers
  if (data[i][linkColumnIndex]) { // Check if the Link cell is not empty
    isLinkColumnEmpty = false;
    break;
  }
}

if (isLinkColumnEmpty) {
  Logger.log("No links found in the Link column.");
  Browser.msgBox("No URLs found in the 'Link' column.", Browser.Buttons.OK);
  return;
}

  Logger.log("Name column index: " + nameColumnIndex);
  Logger.log("Link column index: " + linkColumnIndex);

  // If Form Validation Status column doesn't exist, add it as the last column
  if (statusColumnIndex === -1) {
    statusColumnIndex = headers.length;
    ticketingFormSheet.getRange(1, statusColumnIndex + 1).setValue('Form Validation Status');
  }

//  const exportedDataSheet = ss.getSheetByName('Exported Data from ticketing forms') || ss.insertSheet('Exported Data from ticketing forms');
 // const exportedHeaders = ['Category', 'Sub Categories', 'Form Name', 'Extra Field 1', 'Extra Field 2','Validation Status'];
   
// Check if the 'Exported Data from ticketing forms' sheet exists
let exportedDataSheet = ss.getSheetByName('Exported Data from ticketing forms');
const exportedHeaders = ['Category', 'Sub Categories', 'Form Name', 'Department', 'Extra Field 1', 'Extra Field 2', 'Validation Status'];

if (exportedDataSheet) {
    // Clear the entire sheet including headers and formatting
    exportedDataSheet.clear();
    exportedDataSheet.clearConditionalFormatRules();
    Logger.log("Cleared entire sheet 'Exported Data from ticketing forms'.");
} else {
    // Create a new sheet if it does not exist
    exportedDataSheet = ss.insertSheet('Exported Data from ticketing forms');
    Logger.log("Created new sheet 'Exported Data from ticketing forms'.");
}

// Set new headers
exportedDataSheet.appendRow(exportedHeaders);
Logger.log("Set new headers in 'Exported Data from ticketing forms'.");

// Log completion of the processing step
Logger.log("Exported data sheet processed.");


 
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameColumnIndex];
    const url = data[i][linkColumnIndex];
    const departmentName = data[i][headers.indexOf('Departments')]; // Retrieve the Department Name from Ticketing form sheet

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
        const { validationStatus, validatedData } = validateTicketFormBothCM_1(importSheet, validatedSheetName);
        
        // Update Form Validation Status in TicketingForm sheet
        ticketingFormSheet.getRange(i + 1, statusColumnIndex + 1).setValue(validationStatus);

        // Append validated data to the "Exported Data from ticketing forms" sheet
        appendValidatedData(exportedDataSheet, validatedData, name, departmentName);
        
      } catch (e) {
        Logger.log(`Error processing ${name} with URL ${url}: ${e.toString()}`);
        // Set status to indicate error in case of exception
       
        ticketingFormSheet.getRange(i + 1, statusColumnIndex + 1).setValue('Error: ' + e.toString());
      }
    }
  
  }

  // Add conditional formatting to the exported data sheet
  addConditionalFormatting(exportedDataSheet);
}

function appendValidatedData(exportedDataSheet, validatedData, formName, departmentName) {
  if (validatedData.length === 0) return;

  const headers = exportedDataSheet.getRange(1, 1, 1, exportedDataSheet.getLastColumn()).getValues()[0];
  const categoryIndex = headers.indexOf('Category');
  const subCategoryIndex = headers.indexOf('Sub Categories');
  const formNameIndex = headers.indexOf('Form Name');
  const DepartmentNameIndex = headers.indexOf('Department');
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
    newRow[DepartmentNameIndex] = departmentName;
    newRow[validationStatusIndex] = status;  // Write the status to the new 'Validation Status' column
    exportedDataSheet.appendRow(newRow);
  });
}


function validateTicketFormBothCM_1(importSheet, validatedSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Clear existing validated sheet if it exists
  const existingValidatedSheet = ss.getSheetByName(validatedSheetName);
  if (existingValidatedSheet) {
    ss.deleteSheet(existingValidatedSheet);
    Logger.log("Deleted existing validated sheet: " + validatedSheetName);
  }

  // Create a new validated sheet
  const validatedSheet = ss.insertSheet(validatedSheetName);
  validatedSheet.appendRow(['Category', 'Sub Categories', 'Status']); // Header
  Logger.log("Created new validated sheet: " + validatedSheetName);

  const lastRow = importSheet.getLastRow();

  // Check if the import sheet is blank
  if (lastRow === 0) {
    validatedSheet.appendRow(['No data found in import sheet', 'Error']);
    Logger.log("No data found in import sheet.");
    return { validationStatus: 'No data found', validatedData: [] };
  }

  // Get the header row to dynamically find column indices
  const headerRow = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues()[0];
  const dependsOnColumnIndex = headerRow.indexOf('dependsOnValue') + 1;
  const optionsColumnIndex = headerRow.indexOf('options') + 1;
  const keyColumnIndex = headerRow.indexOf('key') + 1;

  Logger.log("Column indices - dependsOnValue: " + dependsOnColumnIndex + ", options: " + optionsColumnIndex + ", key: " + keyColumnIndex);

  // Get values from the import sheet
  const dependsOnValues = dependsOnColumnIndex > 0 
    ? importSheet.getRange(2, dependsOnColumnIndex, lastRow - 1, 1).getValues().flat()
    : [];
  const keyValues = importSheet.getRange(2, keyColumnIndex, lastRow - 1, 1).getValues().flat();

  Logger.log("Retrieved dependsOnValues: " + dependsOnValues.join(", "));
  Logger.log("Retrieved keyValues: " + keyValues.join(", "));

  // Search for the row index containing the key "category"
  let categoryRowIndex = keyValues.indexOf("category");
  if (categoryRowIndex === -1) {
    categoryRowIndex = 1; // If "category" is not found from row 3, we set it to 1
  }

  // Calculate the actual row number in the spreadsheet where the "category" key is found
  const categoryRow = categoryRowIndex + 2; // Add 2 because we start counting from row 2 in the spreadsheet

  Logger.log("Actual row number for 'category': " + categoryRow);

  // Retrieve options value from the row where "category" key is found
  const options = importSheet.getRange(categoryRow, optionsColumnIndex).getValue();
  const optionsArray = options.split("||").map(option => option.trim());
  const optionsSet = new Set(optionsArray); // Create a set for fast lookup

  Logger.log("Options picked from 'options' column: " + optionsArray.join(", "));

  // Initialize the subCategoryMap object
  const subCategoryMap = {};
  let hasSubCategory = false; // Flag to check if any subcategories are found
  let subCategoryType = ''; // To differentiate between "subCategory" and "category.subCategory.name"

  for (let i = 0; i < keyValues.length; i++) {
    const key = keyValues[i];
    if (key === "subCategory" || key === "category.subCategory.name") {
      hasSubCategory = true; // Mark that a subcategory has been detected
      subCategoryType = key; // Store the type of subcategory key
      const subCategoryValue = dependsOnValues[i];
      if (subCategoryValue) {
        const subCategoryOptions = importSheet.getRange(i + 2, optionsColumnIndex).getValue();
        subCategoryMap[subCategoryValue] = subCategoryOptions; // Store subcategory options as is
        Logger.log("SubCategory found: " + subCategoryValue + " with options: " + subCategoryOptions);
      }
    }
  }

  // Display the appropriate toast message based on the subcategory type detected
  if (hasSubCategory) {
    if (subCategoryType === "subCategory") {
      SpreadsheetApp.getActiveSpreadsheet().toast('New CM form detected: Based on key value "subCategory"', '⚠️ Form Details', 5);
      Logger.log('Toast displayed: New CM form detected: Based on key value "subCategory"');
    } else if (subCategoryType === "category.subCategory.name") {
      SpreadsheetApp.getActiveSpreadsheet().toast('Mono CM form detected: Based on key value "category.subCategory.name"', '⚠️ Form Details', 5);
      Logger.log('Toast displayed: Mono CM form detected: Based on key value "category.subCategory.name"');
    }
  } else {
    // If no subCategory Key found, handle the category only case
    optionsArray.forEach(option => {
      validatedSheet.appendRow([option, '', 'SubCategory Key not available in ticketing form']);
      Logger.log("Option: " + option + " - Status: SubCategory Key not available in ticketing form");
    });
    // Append the validated data for options only to ensure it gets processed
    const validatedData = optionsArray.map(option => [option, '', 'SubCategory Key not available in ticketing form']);
    return { validationStatus: 'No subcategory', validatedData: validatedData };
  }

  let validationStatus = 'Pass'; // Default status
  const validatedData = [];

  // Check the status for each option against its subcategory
  optionsArray.forEach(category => {
    const subCategory = subCategoryMap[category] || ''; // If no subcategory, keep it as empty string
    let status = '';

    if (!subCategory) {
      // If no subcategory is mapped, check if the category is in options
      if (optionsSet.has(category)) {
        if (dependsOnValues.includes(category)) {
          status = 'No Sub category'; // No subcategory associated
        } else {
          status = 'No Sub category key defined'; // Subcategory missing when it should be there
        }
      } else {
        status = 'Missing in Column options; Its critical'; // Category itself is missing in options
      }
      Logger.log("Category: " + category + " - Status: " + status);
    } else {
      // If subcategory is present, check for its presence in the options array
      status = dependsOnValues.includes(category) ? 'Available in Both' : 'Missing in Column options; Its critical';
      Logger.log("Category: " + category + " with Subcategory: " + subCategory + " - Status: " + status);
    }

    validatedSheet.appendRow([category, subCategory, status]);
    validatedData.push([category, subCategory, status]);
     Logger.log("Category: " + category + " - SubCategory: " + subCategory + " - Status: " + status);

  });

  // Check for extra subcategories in dependsOnValue that are not available in options
  dependsOnValues.forEach(dependsOnValue => {
    const subCategoryOptions = subCategoryMap[dependsOnValue];
    if (subCategoryOptions && !optionsSet.has(dependsOnValue)) {
      validatedSheet.appendRow([dependsOnValue, '', 'Missing in Column options; Its critical']);
       validatedData.push([dependsOnValue, '', 'Missing in Column options; Its critical']);
 Logger.log("Extra Subcategory in dependsOnValue: " + dependsOnValue + " - Status: Missing in Column options; Its critical");
    }
  });

return { validationStatus, validatedData };
}

function addConditionalFormatting(sheet) {
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
 
  // Check if sheet has at least one row of data  
  if (sheet.getLastRow() <= 1) {  
  Logger.log("No data to format in 'Exported Data from ticketing forms' sheet");  
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

