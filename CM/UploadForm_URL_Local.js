function uploadLocal_UAT() {
  Upload_ticketing_Form(false, true);
}

function uploadLocal_PROD() {
  Upload_ticketing_Form(true, true);
}

function uploadURL_UAT() {
  Upload_ticketing_Form(false, false);
}

function uploadURL_PROD() {
  Upload_ticketing_Form(true, false);
}


function showUploadFormDialog(isProd) {
  var env = isProd ? 'PROD' : 'UAT';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Upload_TicketingForm');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error', '"Upload_TicketingForm" sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

try {

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var nameIndex = headers.indexOf('Name');
    var formIdIndex = headers.indexOf('Form ID');
    var departmentsIndex = headers.indexOf('Departments');

// in Upload sheet if no data stop further
  if (data.length < 2) {
    Logger.log("No data found in TicketingForm sheet");
    Browser.msgBox("No data found in " + sheet.getSheetName(), Browser.Buttons.OK);
    return;
  }

    if (nameIndex === -1 || formIdIndex === -1 || departmentsIndex === -1) {
      throw new Error('Required columns (Name, Form ID, Departments) not found.');
    }

    var formsData = data.slice(1).map(function(row) {
      return {
        name: row[nameIndex],
        formId: row[formIdIndex],
        departments: row[departmentsIndex]
      };
    });

    var html = HtmlService.createHtmlOutputFromFile('UploadTicketingForm')
      .setWidth(600)
      .setHeight(400);
    
    html.append(`<script>setEnvAndFormData('${env}', ${JSON.stringify(formsData)});</script>`);

    SpreadsheetApp.getUi().showModalDialog(html, 'Upload Ticketing Form from Local System');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error in showUploadFormDialog: ' + e.message);
  }
}


function uploadLocalFile(fileContent, formId, department, fileName, isProd) {
return new Promise((resolve, reject) => {
  try {
    Logger.log('Starting uploadLocalFile function...');
    Logger.log('Parameters received - formId: ' + formId + ', department: ' + department);
    Logger.log('File name: ' + fileName);
    Logger.log('Received isProd value: ' + isProd);

    // Decode base64 and create a blob for fileContent
    var fileBlob = Utilities.newBlob(Utilities.base64Decode(fileContent.split(',')[1]), 'application/octet-stream', fileName);
    Logger.log('File blob created with size: ' + fileBlob.getBytes().length + ' bytes.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('Main');

    if (!mainSheet) {
      Logger.log('Main sheet not found.');
      throw new Error('Main sheet not found.');
    }

    var rowIndex = 2; // Assuming the values are in the 2nd row
    Logger.log('Fetching BOT_ID and JWT based on dashboard domain...');
    
    var BOT_ID, jwt;
    var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];
    var dashboardDomain = mainSheet.getRange(rowIndex, headersValues.indexOf('Dashboard Domain Name') + 1).getValue();
    Logger.log('Dashboard Domain: ' + dashboardDomain);
    
    if (isProd) {
      Logger.log("Using PROD credentials");
      BOT_ID = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD BOT ID') + 1).getValue();
      jwt = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD JWT') + 1).getValue();
      Logger.log('Using PROD credentials: BOT_ID: ' + BOT_ID);
    } else {
      Logger.log("Using UAT credentials");
      BOT_ID = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT BOT ID') + 1).getValue();
      jwt = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT JWT') + 1).getValue();
      Logger.log('Using UAT credentials: BOT_ID: ' + BOT_ID);
    }

    if (!BOT_ID) {
      showProgressToast_tktUpload(ss, 'Error: BOT ID is missing.', 0, 0, 0);
      Logger.log('Error: BOT ID is missing.');
      return { status: 'Error', message: 'BOT ID is missing' };
    }

    if (!jwt) {
      showProgressToast_tktUpload(ss, 'Error: JWT is missing.', 0, 0, 0);
      Logger.log('Error: JWT is missing.');
      return { status: 'Error', message: 'JWT is missing' };
    }

    Logger.log('Uploading file: ' + fileName);

    // Preparing multipart payload
    var payload = {
      'file': fileBlob, // Directly include the fileBlob
      'departments': department,
      'id': formId
    };

    Logger.log('Payload prepared for upload.');

    var headers = {
      'Authorization': jwt,
      'x-cm-dashboard-user': 'true'
    };

    var apiUrl = dashboardDomain + '/@@@@@@@@@@@@@@@@@/' + BOT_ID + '/@@@@@@@@@';
    Logger.log('Sending request to ' + apiUrl);

    var options = {
      method: 'put',
      headers: headers,
      payload: payload, // Pass payload directly
      muteHttpExceptions: true
    };

    // Making the API call
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log('Made API call for Form ID: ' + formId);

    // Process the response
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

    Logger.log('Final Status: ' + status);
    if (errorMessage) {
      Logger.log('Error Message: ' + errorMessage);
    }

    // Update the Upload_TicketingForm sheet
    var sheet = ss.getSheetByName('Upload_TicketingForm');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Error', '"Upload_TicketingForm" sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      Logger.log('"Upload_TicketingForm" sheet not found.');
      return; // Exit the function if the sheet is not found
    }
    Logger.log('Accessed the "Upload_TicketingForm" sheet.');

    // Define column indexes
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var formIdIndex = headers.indexOf('Form ID');
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

    // Add Response Code column if it does not exist
    if (responseCodeIndex === -1) {
      responseCodeIndex = headers.length + 1;
      sheet.getRange(1, responseCodeIndex + 1).setValue('Response Code');
      Logger.log('Added "Response Code" column.');
    }

    // Add Message column if it does not exist
    if (messageIndex === -1) {
      messageIndex = headers.length + 2;
      sheet.getRange(1, messageIndex + 1).setValue('Message');
      Logger.log('Added "Message" column.');
    }

    // Add Response Body column if it does not exist
    if (responseBodyIndex === -1) {
      responseBodyIndex = headers.length + 3;
      sheet.getRange(1, responseBodyIndex + 1).setValue('Response Body');
      Logger.log('Added "Response Body" column.');
    }

    // Re-fetch headers to update index values after adding new columns
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    formIdIndex = headers.indexOf('Form ID');
    statusIndex = headers.indexOf('Status');
    responseCodeIndex = headers.indexOf('Response Code');
    messageIndex = headers.indexOf('Message');
    responseBodyIndex = headers.indexOf('Response Body');

    // Find the row with the matching formId
    var dataRange = sheet.getRange(2, formIdIndex + 1, sheet.getLastRow() - 1, 1);
    var dataValues = dataRange.getValues();
    var targetRow = -1;

    for (var i = 0; i < dataValues.length; i++) {
      if (dataValues[i][0] == formId) {
        targetRow = i + 2; // +2 to adjust for the header row and zero-index
        break;
      }
    }

    if (targetRow > -1) {
      sheet.getRange(targetRow, statusIndex + 1).setValue(status);
      sheet.getRange(targetRow, responseCodeIndex + 1).setValue(responseCode);
      sheet.getRange(targetRow, responseBodyIndex + 1).setValue(truncatedResponseBody);
      sheet.getRange(targetRow, messageIndex + 1).setValue(errorMessage);
      Logger.log('Updated row ' + targetRow + ' with response details.');
    } else {
      Logger.log('Form ID ' + formId + ' not found in the sheet.');
    }

    SpreadsheetApp.flush(); // Force update to the sheet
    Logger.log('Updates flushed to the sheet.');

    var finalResponse = {
      status: status,
      responseBody: truncatedResponseBody,
      errorMessage: errorMessage
    };


Logger.log("Response to be returned to client: " + JSON.stringify(finalResponse));
      resolve(finalResponse); // Resolve the promise with the response

    } catch (error) {
      Logger.log('Error uploading file: ' + error.message);
      reject(error); // Reject the promise with the error
    }
  });
}

function Upload_ticketing_Form(isProd, isLocal = false)  {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  showProgressToast_tktUpload(ss, 'Initializing...', 0, 0, 0);

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
      showProgressToast_tktUpload(ss, 'Error: BOT ID is missing.', 0, 0, 0);
      Logger.log('Error: BOT ID is missing.');
      return;
    }

    if (!jwt) {
      showProgressToast_tktUpload(ss, 'Error: JWT is missing.', 0, 0, 0);
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

// if from Upload Local
Logger.log("Environment selected as PROD " + isProd)
Logger.log("Is Upload through Local "+ isLocal)

    if (isLocal) {
    showUploadFormDialog(isProd);
    return;
  }


  // else continue URL upload
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];
    var rows = data.slice(1);
    var totalRows = rows.length;

    if (totalRows === 0) {
      showProgressToast_tktUpload(ss, 'Nothing to upload.', 0, 0, 0);
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

      showProgressToast_tktUpload(ss, 'Processing...', index + 1, totalRows, totalRows - (index + 1));

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

          var response = UrlFetchApp.fetch(dashboardDomain + '/@@@@@@@@@@@/' + BOT_ID + '/@@@@@@@@@@@@@@@@', {
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
          logLibraryUsage('Upload Form Local Environment PROD = ' + isProd , 'Fail', e.toString());
          }


          sheet.getRange(index + 2, statusIndex + 1).setValue(status);
          sheet.getRange(index + 2, responseCodeIndex + 1).setValue(responseCode);
          sheet.getRange(index + 2, responseBodyIndex + 1).setValue(truncatedResponseBody);
          sheet.getRange(index + 2, messageIndex + 1).setValue(errorMessage);

        // Flush Status forcefully
        SpreadsheetApp.flush();

        // Display toast message with the API response code

        // ss.toast('Processed row ' + (index + 1) + '. API Response Code: ' + responseCode, 'Ticketing Upload Status');
        
        } catch (e) {
          Logger.log('Error for Form ID: ' + formId + ' - ' + e.message);
          sheet.getRange(index + 2, statusIndex + 1).setValue('Error: ' + e.message);
          logLibraryUsage('Upload Form Local Environment PROD = ' + isProd , 'Fail', e.toString());
          SpreadsheetApp.flush();
        }
      } else {
        Logger.log('Skipping row ' + (index + 1) + ' due to missing data.');
        sheet.getRange(index + 2, statusIndex + 1).setValue('Error: Missing data');
      SpreadsheetApp.flush();
      }
       

    });

    showProgressToast_tktUpload(ss, 'Completed', totalRows, totalRows, 0);

  } catch (e) {
    showProgressToast_tktUpload(ss, 'Error: ' + e.message, 0, 0, 0);
    Logger.log('Error: ' + e.message);
  }
}

function showProgressToast_tktUpload(ss, message, processed, total, pending) {
  ss.toast(message + ' Processed: ' + processed + '/' + total + ', Pending: ' + pending, 'Upload Progress', 3);
}
