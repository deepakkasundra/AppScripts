// function SHowProgress is available in COde.js as its common
function getUserEmail() {
  var user_email = Session.getActiveUser().getEmail();
Logger.log(user_email);  
  return user_email ? user_email : null; // Return user email or null if not available
}

function removeFromBlacklistUAT() {
  removeFromBlacklist(false); // UAT Environment
}

function removeFromBlacklistPROD() {
  removeFromBlacklist(true); // PROD Environment
}

function removeFromBlacklist(isProd) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
  var user_email = getUserEmail();

  if (!user_email) {
    Browser.msgBox("Email not found. Please log in again or refresh your browser.", Browser.Buttons.OK);
    return;
  }

  var config;
  try {
    config = getMainSheetData();
  } catch (e) {
    Browser.msgBox("Failed to retrieve configuration from Main sheet.");
    return;
  }

  var jwt = config.flowchatteronJWT;
  var email = isProd ? config.prodEmail : config.uatEmail;
  var domainname = config.flowchatteronDomain;

  if (!jwt || !email || !domainname) {
    var envName = isProd ? "Production" : "UAT";
    Browser.msgBox(envName + " Flow chatteron JWT, Email, or Domain name is missing.", Browser.Buttons.OK);
    return;
  }


//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var mainSheet = ss.getSheetByName('Main');
//   var user_email = getUserEmail();
//   Logger.log(user_email);
//   var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
//   var headersValues = headersRange.getValues()[0];
//   var rowIndex = 2; // Email is available at row 2

//   var jwt, email, domainname;

//   domainname = mainSheet.getRange(rowIndex, headersValues.indexOf('Mono CM Ticketing form') + 1).getValue(); 
// jwt =mainSheet.getRange(rowIndex, headersValues.indexOf('Flow chatteron JWT') + 1).getValue();

//   if (isProd) {
//     email = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD Email') + 1).getValue();
//   } else {
//     email = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT Email') + 1).getValue();
//   }

//   // Check if JWT or Email is missing
//   if (!jwt || !email) {
//     var envName = isProd ? "Production" : "UAT";
//     Browser.msgBox(envName + " JWT or Email is missing.", Browser.Buttons.OK);
//     return;
//   }

  if (!user_email){
    Browser.msgBox("Email not found. Please log in again or refresh your browser.", Browser.Buttons.OK);
    return;
  }

  showProgressToast(ss, 'Removing from Blacklist in ' + (isProd ? 'PROD' : 'UAT') + '...');

  var url = domainname+'/api/blacklist/mail/whitelist';
  Logger.log(url)
  Logger.log(jwt)
  var headers = {
    'Authorization': jwt,
    'Content-Type': 'application/json'
  };


 var payload = JSON.stringify({
    "email": email,
    "whiteListRequestFrom": user_email
  });
Logger.log(payload)

  var options = {
    'method': 'POST',
    'headers': headers,
    'payload': payload,
     'muteHttpExceptions': true
  };

try {
  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();

  // Check if response is successful (200 or 422)
  if (responseCode === 200 || responseCode === 422) {
    showProgressToast(ss, email + ' successfully processed. Response Code: ' + responseCode + ' | Message: ' + responseBody);
    
    // Show a pop-up with response details
    Browser.msgBox('Success', 'Email ' + email + ' processed successfully.\n\n' +
      'Response Code: ' + responseCode + '\n' +
      'Message: ' + responseBody, Browser.Buttons.OK);
  } else {
    Logger.log('❌ Failed to process request. Response Code: ' + responseCode + ' | Body: ' + responseBody);
    
    // Show a pop-up with the error details
    Browser.msgBox('Error', 'Failed to process request.\n\n' +
      'Response Code: ' + responseCode + '\n' +
      'Message: ' + responseBody, Browser.Buttons.OK);
  }
} catch (error) {
  Logger.log('❌ Error processing request: ' + error);
  
  // Show a pop-up with the error message
  Browser.msgBox('Error', 'Error processing request:\n\n' + error.message, Browser.Buttons.OK);
}

}
