function updateFinalResponse() {
  const sheetName = 'Questions';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) return ui.alert(`Sheet "${sheetName}" not found.`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const responseIndex = headers.indexOf('Response');
  if (responseIndex === -1) return ui.alert("Column 'Response' not found.");

  // Create 'Final Response' column if missing
  let finalResponseIndex = headers.indexOf('Final Response');
  if (finalResponseIndex === -1) {
    finalResponseIndex = headers.length;
    sheet.getRange(1, finalResponseIndex + 1).setValue('Final Response');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ui.alert("No data to process.");

  const responses = sheet.getRange(2, responseIndex + 1, lastRow - 1).getValues();

  const finalResponses = responses.map(row => {
    const text = row[0] || "";
    const cleanText = text.split('%----%')[0].trim();
    return [cleanText];
  });

  sheet.getRange(2, finalResponseIndex + 1, finalResponses.length).setValues(finalResponses);

//  ui.alert("Final Response column updated successfully.");
  SpreadsheetApp.getActiveSpreadsheet().toast("Final Response column updated successfully.", "Info", 5);

}

