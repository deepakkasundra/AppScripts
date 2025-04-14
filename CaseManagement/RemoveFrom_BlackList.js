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
  var mainSheet = ss.getSheetByName('Main');
  var user_email = getUserEmail();
  Logger.log(user_email);
  var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
  var headersValues = headersRange.getValues()[0];
  var rowIndex = 2; // Email is available at row 2

  var jwt, email, domainname;

  domainname = mainSheet.getRange(rowIndex, headersValues.indexOf('Mono CM Ticketing form') + 1).getValue(); 
jwt =mainSheet.getRange(rowIndex, headersValues.indexOf('Flow chatteron JWT') + 1).getValue();

  if (isProd) {
    email = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD Email') + 1).getValue();
  } else {
    email = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT Email') + 1).getValue();
  }

  // Check if JWT or Email is missing
  if (!jwt || !email) {
    var envName = isProd ? "Production" : "UAT";
    Browser.msgBox(envName + " JWT or Email is missing.", Browser.Buttons.OK);
    return;
  }

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
    'payload': payload
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode === 200) {
      showProgressToast(ss, email + ' Successfully removed from Blacklist.');
    } else {
      Logger.log('Failed to remove from Blacklist. Response code: ' + responseCode);
      Browser.msgBox('Failed to remove from Blacklist. Response code: ' + responseCode);
    }
  } catch (error) {
    Logger.log('Error removing from Blacklist: ' + error);
    Browser.msgBox('Error removing from Blacklist: ' + error);
  }
}
