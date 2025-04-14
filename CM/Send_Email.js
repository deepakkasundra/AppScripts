function sendEmail() {
  var sheetName = "Triaging"; // Change this to your sheet name
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found.");
    return;
  }
  
  // Get the header row to find column indices
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headers = headerRange.getValues()[0];
  var columnIndex = {}; // Object to store column indices
  
  // Map column indices to header names
  for (var i = 0; i < headers.length; i++) {
    columnIndex[headers[i]] = i + 1; // Column index starts from 1
  }

  // Check if "Email Sent Status" column exists, if not, add it
  if (!columnIndex["Email Sent Status"]) {
    sheet.getRange(1, headers.length + 1).setValue("Email Sent Status");
    columnIndex["Email Sent Status"] = headers.length + 1;
    Logger.log("Added 'Email Sent Status' column.");
  }
  
  // Clear "Email Sent Status" column
  var statusColumnIndex = columnIndex["Email Sent Status"];
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, statusColumnIndex, lastRow - 1).clearContent();
    Logger.log("'Email Sent Status' column cleared.");
  }


  // Assuming data starts from the second row, change if necessary
  var startRow = 2;
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - startRow + 1;
  
  var dataRange = sheet.getRange(startRow, 1, numRows, headers.length); // Use headers.length for dynamic range
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var recordNumber = startRow + i;
    var mailto = row[columnIndex["Email Send to"] - 1] !== undefined ? row[columnIndex["Email Send to"] - 1] : ""; // Check for undefined and set to empty 
   var ccEmails = row[columnIndex["CC Email"] - 1] !== undefined ? row[columnIndex["CC Email"] - 1] : ""; // Check for undefined and set to empty     
    var statusColumnIndex = columnIndex["Email Sent Status"] - 1;
    // Check if mailto is empty, if so, skip processing this row
    if (!mailto) {
      Logger.log("Skipping row " + recordNumber + " because email recipient is empty.");
        sheet.getRange(recordNumber, statusColumnIndex + 1).setValue("No Email ID found, Skipped: ");
      SpreadsheetApp.flush();
      continue;
    }
    
    // Extract other data from the row
    var subject = row[columnIndex["Subject Line"] - 1] !== undefined ? row[columnIndex["Subject Line"] - 1] : ""; 
    var emailBody = row[columnIndex["Email Body"] - 1] !== undefined ? row[columnIndex["Email Body"] - 1] : ""; 
    var expectedCategory = row[columnIndex["Response category 1"] - 1] !== undefined ? row[columnIndex["Response category 1"] - 1] : ""; 
    var expectedSubcategory = row[columnIndex["Response Sub category 1"] - 1] !== undefined ? row[columnIndex["Response Sub category 1"] - 1] : ""; 

    // Compose email
    var message = emailBody + "\n\n" +  "\n\n" + // Extra enter
                  "Expected Category: " + expectedCategory + "\n" +
                  "Expected Subcategory: " + expectedSubcategory;
    
    try {
      // Send email
      MailApp.sendEmail({
        to: mailto,
        cc: ccEmails, // Include CC emails
        subject: subject,
        body: message
      });

      // Log success
    // Update status column with "Success"
      sheet.getRange(recordNumber, statusColumnIndex + 1).setValue("Success");
      SpreadsheetApp.flush();
      SpreadsheetApp.getActiveSpreadsheet().toast('Default sleep time 10 Sec, for sending next Email', 'Email Send for Row : ' +  recordNumber +'', 4);
      Logger.log("Email sent successfully for Row:  " + recordNumber + ", " + subject +"\n\n"+ message);
    } catch (error) {
      // Log error if sending email fails
      Logger.log("Error sending email for Row: " + recordNumber + ", " + subject + ", Error message: " + error +"\n\n"+ message);
    sheet.getRange(recordNumber, statusColumnIndex + 1).setValue("Failed: " + error);
        SpreadsheetApp.flush();
    }

    // Sleep for 10 seconds (10000 milliseconds) before sending the next email
    Utilities.sleep(10000);
  }
      SpreadsheetApp.getActiveSpreadsheet().toast('Send Mail Process Completed', 'Status' , 10);
     
}
