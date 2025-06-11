function Create_ticketing_Form_FromUAT() {
  Create_ticketing_Form(false);
}

function Create_ticketing_Form_FromPROD() {
  Create_ticketing_Form(true);
}

function Create_ticketing_Form(isProd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  showProgressToast_tktUpload(ss, 'Initializing Form Creation...', 0, 0, 0);

  try {
    var mainSheet = ss.getSheetByName('Main');
    if (!mainSheet) {
      throw new Error('Main sheet not found.');
    }

    var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];
    var rowIndex = 2;
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

    if (!BOT_ID || !jwt) {
      throw new Error('BOT ID or JWT is missing.');
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
      showProgressToast_tktUpload(ss, 'Nothing to create.', 0, 0, 0);
      return;
    }


          const endpoints = getApiEndpoints();
// const loginUrl =  + getValidatedEndpoint(endpoints,'ACL Login'); 
Domain Domain =Domain + '/<REDACTED_PATH>/' + Domain + Domain(Domain,'Domain Domain Domain');
// var response = UrlFetchApp.fetch(dashboardDomain + '/bots/' + BOT_ID + '/cm/ticket-form/create-form-excel', {

Logger.log(url);

    // Define column indexes
    var nameIndex = headers.indexOf('Name');
    var departmentsIndex = headers.indexOf('Departments');
    var formFileUrlIndex = headers.indexOf('Form File URL');
    var responseCodeIndex = headers.indexOf('Response Code');
    var statusIndex = headers.indexOf('Status');
    var messageIndex = headers.indexOf('Message');
    var responseBodyIndex = headers.indexOf('Response Body');

    // Add columns if they do not exist and adjust their indexes
    if (responseCodeIndex === -1) {
      responseCodeIndex = headers.length;
      sheet.getRange(1, responseCodeIndex + 1).setValue('Response Code');
    }
    if (statusIndex === -1) {
      statusIndex = headers.length + 1;
      sheet.getRange(1, statusIndex + 1).setValue('Status');
    }
    if (messageIndex === -1) {
      messageIndex = headers.length + 2;
      sheet.getRange(1, messageIndex + 1).setValue('Message');
    }
    if (responseBodyIndex === -1) {
      responseBodyIndex = headers.length + 3;
      sheet.getRange(1, responseBodyIndex + 1).setValue('Response Body');
    }

    // Clear columns
    sheet.getRange(2, statusIndex + 1, sheet.getLastRow() - 1).clearContent();
    sheet.getRange(2, responseCodeIndex + 1, sheet.getLastRow() - 1).clearContent();
    sheet.getRange(2, responseBodyIndex + 1, sheet.getLastRow() - 1).clearContent();
    sheet.getRange(2, messageIndex + 1, sheet.getLastRow() - 1).clearContent();

    rows.forEach((row, index) => {
      var name = row[nameIndex];
      var departments = row[departmentsIndex];
      var formFileUrl = row[formFileUrlIndex];

      if (name && departments && formFileUrl) {
        try {
          var fileIdMatch = formFileUrl.match(/[-\w]{25,}/);
          if (!fileIdMatch) {
            throw new Error('Invalid Google Drive URL: ' + formFileUrl);
          }
          var fileId = fileIdMatch[0];
          var fileBlob = DriveApp.getFileById(fileId).getBlob();

          var payload = {
            'file': fileBlob,
            'name': name,
            'departments': departments
          };

          var headers = {
            'Authorization': jwt,
            'x-cm-dashboard-user': 'true'
          };



          var response = UrlFetchApp.fetch(url, {
            method: 'post',
            headers: headers,
            payload: payload,
            muteHttpExceptions: true
          });

          var responseCode = response.getResponseCode();
          var responseBody = response.getContentText();
          var status = (responseCode === 200) ? 'Created' : 'Fail to Create';

          var maxCellLength = 49900;
          var truncatedResponseBody = responseBody.length > maxCellLength ? responseBody.substring(0, maxCellLength) + '...(truncated)' : responseBody;

          // Extract message from the response body
          var message = '';
          try {
            var jsonResponse = JSON.parse(responseBody);
            if (jsonResponse && jsonResponse.error && jsonResponse.error.message) {
              message = jsonResponse.error.message;
            } else if (jsonResponse && jsonResponse.errors && jsonResponse.errors[0] && jsonResponse.errors[0].message) {
              message = jsonResponse.errors[0].message;
            }
          } catch (e) {
            Logger.log('Error parsing response body for Form ID: ' + nameIndex + ' - ' + e.message);
            message = 'Error parsing message';
          }

          // Update the sheet with response details
          sheet.getRange(index + 2, statusIndex + 1).setValue(status);
          sheet.getRange(index + 2, responseCodeIndex + 1).setValue(responseCode);
          sheet.getRange(index + 2, responseBodyIndex + 1).setValue(truncatedResponseBody);
          sheet.getRange(index + 2, messageIndex + 1).setValue(message);

          SpreadsheetApp.flush();

        } catch (e) {
          Logger.log('Error for Form ID: ' + nameIndex + ' - ' + e.message);
          sheet.getRange(index + 2, statusIndex + 1).setValue('Error: ' + e.message);
          sheet.getRange(index + 2, responseCodeIndex + 1).setValue('N/A');
          sheet.getRange(index + 2, responseBodyIndex + 1).setValue('N/A');
          sheet.getRange(index + 2, messageIndex + 1).setValue('N/A');

          SpreadsheetApp.flush();
        }
      } else {
        Logger.log('Skipping row ' + (index + 1) + ' due to missing Form ID, Departments, or Form File URL.');
        sheet.getRange(index + 2, statusIndex + 1).setValue('Skipped: Missing data');
        sheet.getRange(index + 2, responseCodeIndex + 1).setValue('N/A');
        sheet.getRange(index + 2, responseBodyIndex + 1).setValue('N/A');
        sheet.getRange(index + 2, messageIndex + 1).setValue('N/A');
        SpreadsheetApp.flush();
      }
    });

    showProgressToast_tktUpload(ss, 'Form Creation Completed', totalRows, totalRows, 0);

  } catch (e) {
    showProgressToast_tktUpload(ss, 'Error: ' + e.message, 0, 0, 0);
    Logger.log('Error: ' + e.message);
  }
}

function showProgressToast_tktUpload(ss, message, processed, total, pending) {
  ss.toast(message + ' Processed: ' + processed + '/' + total + ', Pending: ' + pending, 'Upload Progress', 3);
}


