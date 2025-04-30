function logLibraryUsage(functionName, status, errorMessage = '') {
  try {
    // Spreadsheet ID of your main sheet (the one you provided)
    var sheetId = '1Ng1j3AwplEjGIVQIurbTKvlPwcHF5VqWe9p2c19JKR0';

    // Open the spreadsheet and get the "LibraryUsage_report" sheet
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var reportSheet = spreadsheet.getSheetByName('LibraryUsage_report');
    
    // Get the headers in the first row (1st row in the sheet)
    var headers = reportSheet.getRange(1, 1, 1, reportSheet.getLastColumn()).getValues()[0];
    
    // Ensure necessary headers exist (User Email, User Name, Sheet Name, Spreadsheet Name, File Location, Date Time)
    var requiredHeaders = ['User Email', 'User Name', 'Sheet Name', 'Spreadsheet Name', 'File Location', 'Function Name', 'Status', 'Error Message', 'Date Time'];
    var headerMap = {};  // Map to store header positions
    var headersUpdated = false;

    requiredHeaders.forEach(function(header) {
      var colIndex = headers.indexOf(header);
      if (colIndex === -1) {
        // If the header doesn't exist, append it to the end
        reportSheet.getRange(1, headers.length + 1).setValue(header);
        headerMap[header] = headers.length;  // New header column
        headers.push(header);  // Update the headers array
        headersUpdated = true;
      } else {
        headerMap[header] = colIndex;  // Existing header column
      }
    });
    
    // If new headers were added, refresh the headers array
    if (headersUpdated) {
      headers = reportSheet.getRange(1, 1, 1, reportSheet.getLastColumn()).getValues()[0];
    }
    
    // Get the user email (this works if the user has the necessary permissions)
    var userEmail = Session.getActiveUser().getEmail();

    // Get the user name (if available)
    var userName = userEmail.split('@')[0]; // Extracting name before '@'

    // Get the active sheet name (the sheet from which the function is being called)
    var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    
    // Get the spreadsheet name (from the active spreadsheet)
    var spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
    
    // Get the current timestamp
    var timestamp = new Date();

    // Get the file location (URL of the spreadsheet)
    var fileLocation = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getUrl();

    // Prepare the log data dynamically based on available headers
    var rowData = new Array(headers.length).fill('');
    if (headerMap['User Email'] !== undefined) rowData[headerMap['User Email']] = userEmail;
    if (headerMap['User Name'] !== undefined) rowData[headerMap['User Name']] = userName;
    if (headerMap['Sheet Name'] !== undefined) rowData[headerMap['Sheet Name']] = sheetName;
    if (headerMap['Spreadsheet Name'] !== undefined) rowData[headerMap['Spreadsheet Name']] = spreadsheetName;
    if (headerMap['File Location'] !== undefined) rowData[headerMap['File Location']] = fileLocation;
    if (headerMap['Function Name'] !== undefined) rowData[headerMap['Function Name']] = functionName;
    if (headerMap['Status'] !== undefined) rowData[headerMap['Status']] = status;
    if (headerMap['Error Message'] !== undefined) rowData[headerMap['Error Message']] = errorMessage;
    if (headerMap['Date Time'] !== undefined) rowData[headerMap['Date Time']] = timestamp;

    // Find the last row in the sheet and append the data dynamically
    var lastRow = reportSheet.getLastRow();
    reportSheet.getRange(lastRow + 1, 1, 1, headers.length).setValues([rowData]);

  } catch (e) {
    handleError(e);
    Logger.log('Error logging library usage: ' + e.message);
  }
}

