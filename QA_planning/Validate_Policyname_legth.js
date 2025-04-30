function checkPolicyNames() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QA_Report'); // Replace 'YourSheetName' with the name of your sheet.
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get the first row (header row)
    var policyNameColumnIndex = headerRow.indexOf("Policy Name"); // Find the index of the "Policy Name" column

    // Check if the "Policy Name" column exists
    if (policyNameColumnIndex === -1) {
      throw new Error("The 'Policy Name' column was not found in the sheet.");
    }

    var dataRange = sheet.getRange(2, policyNameColumnIndex + 1, sheet.getLastRow() - 1); // Get the range for the "Policy Name" column starting from row 2
    var data = dataRange.getValues();

    // Check if there is any data in the Policy Name column
    if (data.length === 0 || (data.length === 1 && !data[0][0])) {
      var noDataMessage = "No data available in the Policy Name column. Please enter policy names before proceeding.";
      SpreadsheetApp.getActiveSpreadsheet().toast(noDataMessage, "Policy Name Check", 10);
      Browser.msgBox(noDataMessage, Browser.Buttons.OK); // Display a message box with the no data information
      return; // Exit the function if no data is available
    }

    var policyNames = {};
    var duplicateColor = "#FFA500"; // Orange color for duplicate policy names
    var lengthColor = "#FF0000"; // Red color for policy names exceeding the length limit

    var exceededCount = 0;
    var duplicates = [];

    for (var i = 0; i < data.length; i++) {
      var cellValue = data[i][0];
      var policyName = String(cellValue).trim();
      
      // Check for duplicate policy names
      if (policyNames[policyName]) {
        duplicates.push(policyName);
        sheet.getRange(i + 2, policyNameColumnIndex + 1).setBackground(duplicateColor);
      } else {
        policyNames[policyName] = true; // Track unique policy names
      }

      // Check for policy names exceeding the length limit
      if (policyName.length > 100) {
        sheet.getRange(i + 2, policyNameColumnIndex + 1).setBackground(lengthColor);
        exceededCount++;
      }
      // Update the Policy Name in the same cell with trimmed value
      data[i][0] = policyName;
    }
    
    // Set the updated values back to the sheet
    dataRange.setValues(data);

    // Check the count of unique policy names
    var uniquePolicyCount = Object.keys(policyNames).length;

    // Display a toast message if the unique policy name count exceeds 80 and stop execution
    if (uniquePolicyCount > 80) {
      var message = "Note: Google Sheets allows a maximum of 200 sheets. Size Limit: The total size limit for a Google Sheets file is 10 million cells. Depending on the number of columns and rows you use. It is suggested to *Please create the Planning sheet policy report in parts* in 1-80, 81-160 and so on.";
      SpreadsheetApp.getActiveSpreadsheet().toast("Please create the policy report in parts.", "Policy Name Check", 20);
      Browser.msgBox(message, Browser.Buttons.OK); // Display a message box with the detailed information
      // throw new Error("Execution stopped: Unique policy count exceeds limit."); // Stop further execution
    } else {
      // Display a toast message with the count of exceeded policy names and duplicate policy names
      if (exceededCount > 0 && duplicates.length > 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Found " + exceededCount + " Policy Names with length greater than 100 and " + duplicates.length + " Duplicate Policy Names.", "Policy Name Check", 20);
      } else if (exceededCount > 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Found " + exceededCount + " Policy Names with length greater than 100.", "Policy Name Check", 20);
      } else if (duplicates.length > 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Found " + duplicates.length + " Duplicate Policy Names.", "Policy Name Check", 20);
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast("No issues found with Policy Names.", "Policy Name Check", 10);
      }
    }
  }catch (e) {
  handleError(e);
  logLibraryUsage('Old Check Policy Name', 'Fail', e.toString());
 Logger.log("Error: " + e.message); // Log the error for debugging purposes
  }
}

// function checkPolicyNames() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QA_Report'); // Replace 'YourSheetName' with the name of your sheet.
//   var dataRange = sheet.getRange("B2:B" + sheet.getLastRow()); // Assuming your data starts from row 2 in column B.
//   var data = dataRange.getValues();

//   var policyNames = {};
//   var duplicateColor = "#FFA500"; // Orange color for duplicate policy names
//   var lengthColor = "#FF0000"; // Red color for policy names exceeding the length limit

//   var exceededCount = 0;
//   var duplicates = [];

//   for (var i = 0; i < data.length; i++) {
//     var cellValue = data[i][0];
//     var policyName = String(cellValue).trim();
    
//     // Check for duplicate policy names
//     if (policyNames[policyName]) {
//       duplicates.push(policyName);
//       sheet.getRange(i + 2, 2).setBackground(duplicateColor);
//     } else {
//       policyNames[policyName] = true;
//     }

//     // Check for policy names exceeding the length limit
//     if (policyName.length > 100) {
//       sheet.getRange(i + 2, 2).setBackground(lengthColor);
//       exceededCount++;
//     }
//   // Update the Policy Name in the same cell with trimmed value
//     data[i][0] = policyName;
//   }
//   // Set the updated values back to the sheet
//   dataRange.setValues(data);

//   // Display a toast message with the count of exceeded policy names and duplicate policy names
//   if (exceededCount > 0 && duplicates.length > 0) {
//     SpreadsheetApp.getActiveSpreadsheet().toast("Found " + exceededCount + " Policy Names with length greater than 100 and " + duplicates.length + " Duplicate Policy Names.", "Policy Name Check", 20);
//   } else if (exceededCount > 0) {
//     SpreadsheetApp.getActiveSpreadsheet().toast("Found " + exceededCount + " Policy Names with length greater than 100.", "Policy Name Check", 20);
//   } else if (duplicates.length > 0) {
//     SpreadsheetApp.getActiveSpreadsheet().toast("Found " + duplicates.length + " Duplicate Policy Names.", "Policy Name Check", 20);
//   } else {
//     SpreadsheetApp.getActiveSpreadsheet().toast("No issues found with Policy Names.", "Policy Name Check", 10);
//   }
// }

