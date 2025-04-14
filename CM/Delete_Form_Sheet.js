function deleteSheetsBasedOnNames() {
  const sheetName = "TicketingForm";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found.`);
    ss.toast(`Sheet "${sheetName}" not found.`, 'Error', 5);
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  if (data.length <= 1) {
    ss.toast('No data found in TicketingForm to process.', 'No Data', 5);
    return;
  }
  
  const headers = data[0];

  // Find the index of the "Name" and "Deletion Status" columns
  const nameColumnIndex = headers.indexOf("Name");
  let statusColumnIndex = headers.indexOf("Deletion Status");

  // If "Deletion Status" column doesn't exist, add it
  if (statusColumnIndex === -1) {
    statusColumnIndex = headers.length;
    sheet.getRange(1, statusColumnIndex + 1).setValue("Deletion Status");
  }

  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameColumnIndex];
    if (!name) continue;

    let deletionStatus = "";

    try {
      // Delete the main sheet
      const mainSheet = ss.getSheetByName(name);
      if (mainSheet) {
        ss.deleteSheet(mainSheet);
        deletionStatus += "Deleted main sheet; ";
      } else {
        deletionStatus += "Main sheet not found; ";
      }

      // Delete the validated sheet
      const validatedSheet = ss.getSheetByName(name + " validated");
      if (validatedSheet) {
        ss.deleteSheet(validatedSheet);
        deletionStatus += "Deleted validated sheet; ";
      } else {
        deletionStatus += "Validated sheet not found; ";
      }

      ss.toast(`Processing ${name}`, 'In Progress', 3);

    } catch (e) {
      Logger.log(`Error deleting sheet "${name}" or "${name} validated": ${e.message}`);
      deletionStatus = `Error: ${e.message}`;
      ss.toast(`Error deleting sheet "${name}" or "${name} validated": ${e.message}`, 'Error', 5);
    }

    // Update the deletion status
    sheet.getRange(i + 1, statusColumnIndex + 1).setValue(deletionStatus);
  }

  ss.toast('Sheet deletion process completed.', 'Task Completed', 5);
}

function listAllSheetNames() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get all sheets
    var sheets = ss.getSheets();

    // Check if 'Available_sheet' exists, if not create it
    var availableSheet = ss.getSheetByName("Available_sheet");
    if (!availableSheet) {
      availableSheet = ss.insertSheet("Available_sheet");
    } else {
      // Clear the existing content
      availableSheet.clear();
    }

    // Set the header for the sheet names
    availableSheet.getRange("A1").setValue("Sheet Names");

    // List all sheet names in the Available_sheet
    for (var i = 0; i < sheets.length; i++) {
      availableSheet.getRange(i + 2, 1).setValue(sheets[i].getName());
    }

    ss.toast('Sheet names have been successfully listed.', 'Task Completed', 5);

  } catch (e) {
handleError(e);
    Logger.log('Error listing sheet names: ' + e.message);
    ss.toast('Error listing sheet names: ' + e.message, 'Error', 5);
  }
}
