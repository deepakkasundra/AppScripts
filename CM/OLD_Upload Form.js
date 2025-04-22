function OLD_Upload_ticketing_FromUAT() {
  OLD_Upload_ticketing_Form(false);
}

function OLD_Upload_ticketing_FromPROD() {
  OLD_Upload_ticketing_Form(true);
}

function OLD_Upload_ticketing_Form(isProd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  showProgressToast_tktUpload_OLD(ss, 'Initializing...', 0, 0, 0);

  try {
    var mainSheet = ss.getSheetByName('Main');
    if (!mainSheet) {
      throw new Error('Main sheet not found.');
    }

    var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];
    var rowIndex = 2; // as value available at 2
    var dashboardDomain = mainSheet.getRange(rowIndex, headersValues.indexOf('Dashboard Domain Name') + 1).getValue();

    var BOT_ID, jwt;

    if (isProd) {
      BOT_ID = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD BOT ID') + 1).getValue();
      jwt = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD JWT') + 1).getValue();
      Logger.log('Using PROD credentials: BOT_ID: ' + BOT_ID + ', JWT: ' + jwt);
    } else {
      BOT_ID = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT BOT ID') + 1).getValue();
      jwt = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT JWT') + 1).getValue();
      Logger.log('Using UAT credentials: BOT_ID: ' + BOT_ID + ', JWT: ' + jwt);
    }

    if (!BOT_ID) {
      showProgressToast_tktUpload_OLD(ss, 'Error: BOT ID is missing.', 0, 0, 0);
      Logger.log('Error: BOT ID is missing.');
      return;
    }

    if (!jwt) {
      showProgressToast_tktUpload_OLD(ss, 'Error: JWT is missing.', 0, 0, 0);
      Logger.log('Error: JWT is missing.');
      return;
    }

    var sheet = ss.getSheetByName('Upload_TicketingForm');
    if (!sheet) {
      //      throw new Error('"Upload_TicketingForm" sheet not found.');
            SpreadsheetApp.getUi().alert('Error', '"Upload_TicketingForm" sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
            return; // Exit the function if the sheet is not found
    }
    Logger.log('Accessed the "Upload_TicketingForm" sheet.');

    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];
    var rows = data.slice(1);
    var totalRows = rows.length;

    if (totalRows === 0) {
      showProgressToast_tktUpload_OLD(ss, 'Nothing to upload.', 0, 0, 0);
      Logger.log('Nothing to upload.');
      return;
    }

    // Define column indexes
    var formIdIndex = headers.indexOf('Form ID');
    var nameIndex = headers.indexOf('Name');
    var departmentsIndex = headers.indexOf('Departments');
    var formFileUrlIndex = headers.indexOf('Form File URL');
    var statusIndex = headers.indexOf('Status');
    var responseCodeIndex = headers.indexOf('Response Code');
    var messageIndex = headers.indexOf('Message');
    var responseBodyIndex = headers.indexOf('Response Body');

    // Add Status column if it does not exist
    if (statusIndex === -1) {
      statusIndex = headers.length;
      sheet.getRange(1, statusIndex + 1).setValue('Status');
      Logger.log('Added "Status" column.');
    }

    // Add Response Code column if it does not exist and adjust index if needed
    if (responseCodeIndex === -1) {
      responseCodeIndex = headers.length + 1;
      sheet.getRange(1, responseCodeIndex + 1).setValue('Response Code');
      Logger.log('Added "Response Code" column.');
    }

    // Add Message column if it does not exist and adjust index if needed
    if (messageIndex === -1) {
      messageIndex = headers.length + 2;
      sheet.getRange(1, messageIndex + 1).setValue('Message');
      Logger.log('Added "Message" column.');
    }
    
    // Add Response Body column if it does not exist and adjust index if needed
    if (responseBodyIndex === -1) {
      responseBodyIndex = headers.length + 3;
      sheet.getRange(1, responseBodyIndex + 1).setValue('Response Body');
      Logger.log('Added "Response Body" column.');
    }

    // Clear Status, Response Code, Message, and Response Body columns
    if (statusIndex !== -1) {
      sheet.getRange(2, statusIndex + 1, sheet.getLastRow() - 1).clearContent();
      Logger.log('Cleared "Status" column.');
    }

    if (responseCodeIndex !== -1) {
      sheet.getRange(2, responseCodeIndex + 1, sheet.getLastRow() - 1).clearContent();
      Logger.log('Cleared "Response Code" column.');
    }

    if (messageIndex !== -1) {
      sheet.getRange(2, messageIndex + 1, sheet.getLastRow() - 1).clearContent();
      Logger.log('Cleared "Message" column.');
    }

    if (responseBodyIndex !== -1) {
      sheet.getRange(2, responseBodyIndex + 1, sheet.getLastRow() - 1).clearContent();
      Logger.log('Cleared "Response Body" column.');
    }


    rows.forEach((row, index) => {
      var formId = row[formIdIndex];
      var name = row[nameIndex];
      var departments = row[departmentsIndex];
      var formFileUrl = row[formFileUrlIndex];

      Logger.log('Processing row ' + (index + 1) + ': Form ID: ' + formId + ', Name: ' + name + ', Departments: ' + departments + ', Form File URL: ' + formFileUrl);

      showProgressToast_tktUpload_OLD(ss, 'Processing...', index + 1, totalRows, totalRows - (index + 1));

      if (formId && departments && formFileUrl) {
        try {
          var fileIdMatch = formFileUrl.match(/[-\w]{25,}/);
          if (!fileIdMatch) {
            throw new Error('Invalid Google Drive URL: ' + formFileUrl);
          }
          var fileId = fileIdMatch[0];
          Logger.log('Extracted file ID: ' + fileId);

          var fileBlob = DriveApp.getFileById(fileId).getBlob();
          Logger.log('Fetched file blob for file ID: ' + fileId);

          var payload = {
            'file': fileBlob,
            'departments': departments,
            'id': formId
          };
          Logger.log('Defined payload for Form ID: ' + formId);

          var headers = {
            'Authorization': jwt,
            'x-cm-dashboard-user': 'true'
          };
          Logger.log('Defined headers for API call.');

          var response = UrlFetchApp.fetch(dashboardDomain + '/@@@@@@@@@@@/' + BOT_ID + '/@@@@@@@@@@@@', {
            method: 'put',
            headers: headers,
            payload: payload,
            muteHttpExceptions: true
          });
          Logger.log('Made API call for Form ID: ' + formId);

          var responseCode = response.getResponseCode();
          var responseBody = response.getContentText();
          Logger.log('Response Code: ' + responseCode);
          Logger.log('Response Body: ' + responseBody);

          var status = (responseCode === 200) ? 'Uploaded' : 'Fail to Upload';
          var maxCellLength = 49900;
          var truncatedResponseBody = responseBody.length > maxCellLength ? responseBody.substring(0, maxCellLength) + '...(truncated)' : responseBody;

          var errorMessage = '';
          try {
            var jsonResponse = JSON.parse(responseBody);
            if (jsonResponse.errors && jsonResponse.errors.length > 0) {
              errorMessage = jsonResponse.errors[0].message;
            } else if (jsonResponse.error && jsonResponse.error.message) {
              errorMessage = jsonResponse.error.message;
            }
          } catch (e) {
            errorMessage = 'Error parsing response: ' + e.message;
          }


          sheet.getRange(index + 2, statusIndex + 1).setValue(status);
          sheet.getRange(index + 2, responseCodeIndex + 1).setValue(responseCode);
          sheet.getRange(index + 2, responseBodyIndex + 1).setValue(truncatedResponseBody);
          sheet.getRange(index + 2, messageIndex + 1).setValue(errorMessage);

        // Flush Status forcefully
        SpreadsheetApp.flush();

        } catch (e) {
          Logger.log('Error for Form ID: ' + formId + ' - ' + e.message);
          sheet.getRange(index + 2, statusIndex + 1).setValue('Error: ' + e.message);
          SpreadsheetApp.flush();
        }
      } else {
        Logger.log('Skipping row ' + (index + 1) + ' due to missing data.');
        sheet.getRange(index + 2, statusIndex + 1).setValue('Error: Missing data');
      SpreadsheetApp.flush();
      }
       

    });

    showProgressToast_tktUpload_OLD(ss, 'Completed', totalRows, totalRows, 0);

  } catch (e) {
    showProgressToast_tktUpload_OLD(ss, 'Error: ' + e.message, 0, 0, 0);
    Logger.log('Error: ' + e.message);
  }
}

function showProgressToast_tktUpload_OLD(ss, message, processed, total, pending) {
  ss.toast(message + ' Processed: ' + processed + '/' + total + ', Pending: ' + pending, 'Upload Progress', 3);
}
