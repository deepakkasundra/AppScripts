function consolidateFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Seprated Data");
  const resultSheetName = "Form_wise_SeparatedData";

  // Check if the result sheet already exists; if not, create it
  let resultSheet = ss.getSheetByName(resultSheetName);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
  } else {
    resultSheet.clear(); // Clear existing data if the sheet already exists
  }

  // Set headers for the result sheet
  resultSheet.getRange(1, 1, 1, 3).setValues([["Category", "Sub Categories", "Form Name"]]);

  // Get all the data from the "Separated Data" sheet
  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove headers

  // Create a map to store consolidated data
  const consolidatedData = {};

  // Find the column indices for "Category", "Sub Categories", and "Form Name"
  const categoryIndex = headers.indexOf("Category");
  const subCategoryIndex = headers.indexOf("Sub Categories");
  const formNameIndex = headers.indexOf("Form Name");

  // Consolidate data
  data.forEach(row => {
    const category = row[categoryIndex];
    const subCategory = row[subCategoryIndex];
    const formName = row[formNameIndex];
    const key = category + "|" + subCategory;

    if (!consolidatedData[key]) {
      consolidatedData[key] = { category, subCategory, formNames: [] };
    }
    if (!consolidatedData[key].formNames.includes(formName)) {
      consolidatedData[key].formNames.push(formName);
    }
  });

  // Prepare data to write to the result sheet
  const resultData = Object.values(consolidatedData).map(item => {
    return [item.category, item.subCategory, item.formNames.join(", ")];
  });

  // Write the consolidated data to the result sheet
  resultSheet.getRange(2, 1, resultData.length, 3).setValues(resultData);
}

