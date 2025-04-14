function verify_ticketing_form_both_CM() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheetName = 'Form';
  var newSheetName = 'Ticketing from verification';

  var originalSheet = spreadsheet.getSheetByName(originalSheetName);
  var lastRow = originalSheet.getLastRow();

  // Check if the "Form" sheet is blank
  if (lastRow == 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Form sheet is blank", "Error", 5);
    Logger.log("Form sheet is blank.");
    return;
  }

  // Clear the existing "Ticketing from verification" sheet if it exists
  var existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
    Logger.log("Deleted existing sheet: " + newSheetName);
  }

  // Create a new "Ticketing from verification" sheet
  var newSheet = spreadsheet.insertSheet(newSheetName);
  newSheet.appendRow(['Category', 'Sub categories', 'Status']); // Header
  Logger.log("Created new sheet: " + newSheetName);

  // Get the header row to dynamically find column indices
  var headerRow = originalSheet.getRange(1, 1, 1, originalSheet.getLastColumn()).getValues()[0];
  var dependsOnColumnIndex = headerRow.indexOf('dependsOnValue') + 1;
  var optionsColumnIndex = headerRow.indexOf('options') + 1;
  var keyColumnIndex = headerRow.indexOf('key') + 1; // Add the index for the "key" column

  Logger.log("Column indices - dependsOnValue: " + dependsOnColumnIndex + ", options: " + optionsColumnIndex + ", key: " + keyColumnIndex);

  // Check if "dependsOnValue" column is available
  var dependsOnValues = [];
  if (dependsOnColumnIndex > 0) {
    dependsOnValues = originalSheet.getRange(2, dependsOnColumnIndex, lastRow - 1, 1).getValues().flat();
    Logger.log("Retrieved dependsOnValues: " + dependsOnValues.join(", "));
  }

  var keyValues = originalSheet.getRange(2, keyColumnIndex, lastRow - 1, 1).getValues().flat(); // Get key values
  Logger.log("Retrieved keyValues: " + keyValues.join(", "));

  // Search for the row index containing the key "category"
  var categoryRowIndex = keyValues.indexOf("category");
  if (categoryRowIndex === -1) {
    categoryRowIndex = 1; // If "category" is not found from row 3, we set it to 1
  }
  Logger.log("Category key found at row index: " + categoryRowIndex);

  // Calculate the actual row number in the spreadsheet where the "category" key is found
  var categoryRow = categoryRowIndex + 2; // Add 2 because we start counting from row 2 in the spreadsheet
  Logger.log("Actual row number for 'category': " + categoryRow);

  // Retrieve options value from the row where "category" key is found
  var options = originalSheet.getRange(categoryRow, optionsColumnIndex).getValue();
  var optionsArray = options.split("||").map(option => option.trim());
  var optionsSet = new Set(optionsArray); // Create a set for fast lookup
  Logger.log("Options picked from 'options' column: " + optionsArray.join(", "));

  // Initialize the subCategoryMap object
  var subCategoryMap = {};

  // Create a mapping of subcategories to categories (options)
  var hasSubCategory = false;
  var subCategoryType = ''; // To differentiate between "subCategory" and "category.subCategory.name"

  for (var i = 0; i < keyValues.length; i++) {
    var key = keyValues[i];
    if (key === "subCategory" || key === "category.subCategory.name") {
      hasSubCategory = true; // Mark that a subcategory has been detected
      subCategoryType = key; // Store the type of subcategory key
      var subCategoryValue = dependsOnValues[i];
      if (subCategoryValue) {
        var subCategoryOptions = originalSheet.getRange(i + 2, optionsColumnIndex).getValue();
        subCategoryMap[subCategoryValue] = subCategoryOptions; // Store subcategory options as is
        Logger.log("SubCategory found: " + subCategoryValue + " with options: " + subCategoryOptions);
      }
    }
  }

  // Display the appropriate toast message based on the subcategory type detected
  if (hasSubCategory) {
    if (subCategoryType === "subCategory") {
      SpreadsheetApp.getActiveSpreadsheet().toast('New CM form detected: Based on key value "subCategory"', '⚠️ Form Details', 20);
      Logger.log('Toast displayed: New CM form detected: Based on key value "subCategory"');
    } else if (subCategoryType === "category.subCategory.name") {
      SpreadsheetApp.getActiveSpreadsheet().toast('Mono CM form detected: Based on key value "category.subCategory.name"', '⚠️ Form Details', 20);
      Logger.log('Toast displayed: Mono CM form detected: Based on key value "category.subCategory.name"');
    }
  }

  // Check the status for each value where the key matches "category.subCategory.name" or "subCategory"
  for (var i = 0; i < optionsArray.length; i++) {
    var category = optionsArray[i];
    var subCategory = subCategoryMap[category] || ''; // If no subcategory, keep it as empty string
    var status = '';

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

    newSheet.appendRow([category, subCategory, status]);
  }

  // Check for extra subcategories in dependsOnValue that are not available in options
  dependsOnValues.forEach(dependsOnValue => {
    var subCategoryOptions = subCategoryMap[dependsOnValue];
    if (subCategoryOptions && !optionsSet.has(dependsOnValue)) {
      newSheet.appendRow([dependsOnValue, '', 'Missing in Column options; Its critical']);
      Logger.log("Extra Subcategory in dependsOnValue: " + dependsOnValue + " - Status: Missing in Column options; Its critical");
    }
  });
}
