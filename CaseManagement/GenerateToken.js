    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Main'); // Change 'Main Sheet' to your sheet name
    var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];
    // var nlpTokenColumnIndex = headersValues.indexOf('NLP Token') + 1; // Adding 1 to convert from 0-based index to 1-based index
    var rowIndex = 2; // as value availabe at 2
    var uatEmail = sheet.getRange(rowIndex, headersValues.indexOf('UAT Email') + 1).getValue();
    var prodEmail = sheet.getRange(rowIndex, headersValues.indexOf('PROD Email') + 1).getValue();
    var password = sheet.getRange(rowIndex, headersValues.indexOf('Password') + 1).getValue();
    var ACL_Domain = sheet.getRange(rowIndex, headersValues.indexOf('ACL Domain Names') + 1).getValue();

function loginAndStoreToken() {

  try {

//    logLibraryUsage('Generate Token');

if (!prodEmail && !uatEmail) {
//  SpreadsheetApp.getActiveSpreadsheet().toast('ProdEmail and uatEmail are blank. Script stopped.');
        Browser.msgBox('Prod Email and UAT Email are blank. Script stopped.', Browser.Buttons.OK);
return;
}
    // UAT Login
    var uatToken, uatActiveBot;
    if (uatEmail) {
      try {
        uatToken = loginAndGetToken(uatEmail, password);
        if (uatToken) {
          uatActiveBot = getActiveBot(uatToken);
      sheet.getRange(rowIndex, headersValues.indexOf('UAT BOT ID') + 1).setValue(uatActiveBot);
      sheet.getRange(rowIndex, headersValues.indexOf('UAT JWT') + 1).setValue('JWT ' + uatToken);
   logLibraryUsage('Generate UAT Token', 'Pass');  // Log UAT success
       //   mainSheet.getRange('D2').setValue(uatActiveBot);
         // mainSheet.getRange('E2').setValue('JWT ' + uatToken);
        }
      } catch (error) {
        logLibraryUsage('Generate UAT Token', 'Fail', error.toString());  // Log UAT failure
        Browser.msgBox(error.toString(), Browser.Buttons.OK);
        SpreadsheetApp.getActiveSpreadsheet().toast('Error during UAT login: ' + error.toString());
      }
    }

    // Prod Login
    var prodToken, prodActiveBot;
    if (prodEmail) {
      try {
        prodToken = loginAndGetToken(prodEmail, password);

        if (prodToken) {
          prodActiveBot = getActiveBot(prodToken);
   sheet.getRange(rowIndex, headersValues.indexOf('PROD BOT ID') + 1).setValue(prodActiveBot);
      sheet.getRange(rowIndex, headersValues.indexOf('PROD JWT') + 1).setValue('JWT ' + prodToken);
  
 logLibraryUsage('Generate PROD Token', 'Pass');
   //       mainSheet.getRange('F2').setValue(prodActiveBot);
     //     mainSheet.getRange('G2').setValue('JWT ' + prodToken);
        }
      } catch (error) {
    logLibraryUsage('Generate PROD Token', 'Fail', error.toString());  // Log PROD failure        SpreadsheetApp.getActiveSpreadsheet().toast('Error during Prod login: ' + error.toString());
      }
    }
  } catch (error) {
    // Catch any errors
        logLibraryUsage('Generate Token', 'Fail', error.toString());  // Log PROD failure
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + error.toString());
  }
}


function loginAndGetToken(email, password) {


  var loginUrl = ACL_Domain + '/api/users/login';
  Logger.log(loginUrl);
  var loginHeaders = {
    'Content-Type': 'application/json'
  };
  
  var loginPayload = {
    'email': email,
    'password': password
  };

  var loginResponse = UrlFetchApp.fetch(loginUrl, {
    'method': 'post',
    'headers': loginHeaders,
    'payload': JSON.stringify(loginPayload),
    'muteHttpExceptions': true
  });

  if (loginResponse.getResponseCode() === 200) {
Logger.log(loginResponse)
    var loginData = JSON.parse(loginResponse.getContentText());
    return loginData.token;
  } else {
    throw new Error('Error logging in: ' + loginResponse.getResponseCode() + ' - ' + loginResponse.getContentText());
  }
}

function getActiveBot(token) {
  var dashboardInitUrl = ACL_Domain + '/api/users/dashboard-init';
  var dashboardInitHeaders = {
    'Authorization': 'JWT ' + token,
    'Content-Type': 'application/json'
  };

  var dashboardInitResponse = UrlFetchApp.fetch(dashboardInitUrl, {
    'method': 'post',
    'headers': dashboardInitHeaders,
    'muteHttpExceptions': true
  });

  if (dashboardInitResponse.getResponseCode() === 200) {
    var dashboardInitData = JSON.parse(dashboardInitResponse.getContentText());
    return dashboardInitData.activeBot;
  } else {
    throw new Error('Error fetching activeBot: ' + dashboardInitResponse.getResponseCode() + ' - ' + dashboardInitResponse.getContentText());
  }
}
