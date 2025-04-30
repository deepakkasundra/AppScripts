function CopyWotkLM_Autonomous_Template() {
  try {
    // Source and destination workbook IDs
    const sourceSpreadsheetId = "1WuPm9X07AMu9bHOGCjUavgYiuafPBAdBtHk-wAEJZ5A";
    const destinationSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    
    // Sheets to copy
    const sheetsToCopy = ["WorkLM_Test Cases Format", "Autonomous_Test Cases Format"];
    
    // Open workbooks
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetId);
    
    // Check existing sheets in the destination
    const existingSheets = destinationSpreadsheet.getSheets().map(sheet => sheet.getName());
    
    sheetsToCopy.forEach(sheetName => {
      const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
      if (!sourceSheet) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Sheet "${sheetName}" does not exist in the source workbook.`, "Warning", 5);
        return;
      }
      
      if (existingSheets.includes(sheetName)) {
        // Prompt user for confirmation
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
          `Sheet "${sheetName}" already exists.`,
          `Do you want to replace it with a new copy from the source workbook?`,
          ui.ButtonSet.YES_NO
        );
        
        if (response === ui.Button.YES) {
          // Delete the existing sheet
          const sheetToDelete = destinationSpreadsheet.getSheetByName(sheetName);
          destinationSpreadsheet.deleteSheet(sheetToDelete);
          
          // Copy the new sheet
          sourceSheet.copyTo(destinationSpreadsheet).setName(sheetName);
          SpreadsheetApp.getActiveSpreadsheet().toast(`Sheet "${sheetName}" has been replaced.`, "Success", 5);
        } else if (response === ui.Button.NO) {
          SpreadsheetApp.getActiveSpreadsheet().toast(`Sheet "${sheetName}" was not replaced.`, "Info", 5);
        } else {
          SpreadsheetApp.getActiveSpreadsheet().toast(`Operation canceled for sheet "${sheetName}".`, "Info", 5);
        }
      } else {
        // Copy the sheet if it doesn't exist
        sourceSheet.copyTo(destinationSpreadsheet).setName(sheetName);
        SpreadsheetApp.getActiveSpreadsheet().toast(`Sheet "${sheetName}" has been copied.`, "Success", 5);
      }
    });
  } catch (error) {
    // Error handling
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, "Error", 5);
  }
}

