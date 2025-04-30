function Seprate_category_subcategory_Employeecode_Extrafields_Optimized() {
try{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var finalSheet = spreadsheet.getSheetByName('Seprated Data');

  if (finalSheet) {
 
 var lastColumn = finalSheet.getLastColumn();
  if (lastColumn > 0) {
    finalSheet.getRange(1, 1, 1, lastColumn).clearComment(); // Clears comments in the header row
  }
    finalSheet.clear();


  } else {
    finalSheet = spreadsheet.insertSheet("Seprated Data");
  }

  var sourceSheet = spreadsheet.getSheetByName('Exported Data from ticketing forms');

  if (!sourceSheet) {
    Logger.log('Source sheet not found: Exported Data from ticketing forms');
    return;
  }

  var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  var categoryIndex = headers.indexOf('Category') + 1;
  var subCategoryIndex = headers.indexOf('Sub Categories') + 1;
  var employeeCodeIndex = headers.indexOf('Extra Field 1') + 1;
  var DepartmentFieldIndex = headers.indexOf('Department') + 1;
  var formNameIndex = headers.indexOf('Form Name') + 1;

  if (categoryIndex === 0 || subCategoryIndex === 0 || employeeCodeIndex === 0 || DepartmentFieldIndex === 0 || formNameIndex === 0) {
    Logger.log('One or more headers not found in the source sheet.');
    return;
  }

//  finalSheet.getRange(1, 1, 1, 5).setValues([['Category', 'Sub Categories', 'Department', 'Extra Field 1', 'Extra Field 2']]);
spreadsheet.setActiveSheet(finalSheet);
var headerValues = [['Category', 'Sub Categories', 'Form Name', 'Department','Extra Field 1']];
var headerRange = finalSheet.getRange(1, 1, 1, headerValues[0].length);
headerRange.setValues(headerValues);

// Add comment to the cell for 'Department'
//var departmentCell = headerRange.getCell(1, 3); // 1st row, 3rd column
//departmentCell.setComment("This column represents the Form Name.");



    // Dynamically find the "Form Name" column index and add a comment to the "Department" cell
    var formNameColumnIndex = headerValues[0].indexOf('Form Name') + 1;
    var departmentColumnIndex = headerValues[0].indexOf('Department') + 1;

    if (formNameColumnIndex > 0) {
      var formnamecell = headerRange.getCell(1, formNameColumnIndex);
      formnamecell.setComment("This column is represents the Form Name.");
    }

    if (departmentColumnIndex > 0) {
      var departmentCell = headerRange.getCell(1, departmentColumnIndex);
      departmentCell.setComment("Here you need yo Update Department Name i.e. Mapped with Ticketing form or provided in SLA");
    }



  var dataRange = sourceSheet.getDataRange();
  var values = dataRange.getValues();

  var finalData = [];

  for (var i = 1; i < values.length; i++) {
    var category = String(values[i][categoryIndex - 1]).split("||").map(function(value) {
      return value.trim();
    });

    var subCategoryValues = String(values[i][subCategoryIndex - 1]).split("||").map(function(value) {
      return value.trim();
    });

    var formName = values[i][formNameIndex - 1]; // Fetching the form name from the same row

    var employeeCodeValues = String(values[i][employeeCodeIndex - 1]).split("||").map(function(value) {
      return value.trim();
    });

    var departmentfield = String(values[i][DepartmentFieldIndex - 1]).split("||").map(function(value) {
      return value.trim();
    });

    for (var a = 0; a < category.length; a++) {
      for (var j = 0; j < subCategoryValues.length; j++) {
        for (var k = 0; k < employeeCodeValues.length; k++) {
          for (var l = 0; l < departmentfield.length; l++) {
            finalData.push([category[a], subCategoryValues[j], formName, departmentfield[l],employeeCodeValues[k] ]);
          }
        }
      }
    }
  }

  if (finalData.length > 0) {
    // Batch write to final sheet
    var batchSize = 1000; // Adjust as needed
    for (var m = 0; m < finalData.length; m += batchSize) {
      finalSheet.getRange(finalSheet.getLastRow() + 1, 1, Math.min(batchSize, finalData.length - m), 5).setValues(finalData.slice(m, m + batchSize));
    }
  }
}catch (e) {
handleError(e);
}
}

