function loginAndStoreToken() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    Logger.log("🟢 Started loginAndStoreToken");

    const mainData = getMainSheetData();
    const sheet = mainData.sheet;
    const rowIndex = mainData.rowIndex;

    const prodEmail = mainData.prodEmail;
    const uatEmail = mainData.uatEmail;
    const password = mainData.password;
    const ACL_Domain = mainData.ACL_Domain;

    if (!prodEmail && !uatEmail) {
      const msg = '❌ Prod Email and UAT Email are blank. Script stopped.';
      Logger.log(msg);
      Browser.msgBox(msg);
      return;
    }

    if (!password || !ACL_Domain) {
      const msg = '❌ Password or ACL Domain Name is missing. Script stopped.';
      Logger.log(msg);
      Browser.msgBox(msg);
      return;
    }

    // UAT Login
    if (uatEmail) {
      try {
        const uatToken = loginAndGetToken(uatEmail, password, ACL_Domain);
        if (uatToken) {
          const uatActiveBot = getActiveBot(uatToken, ACL_Domain);
          sheet.getRange(rowIndex, mainData.headers.indexOf('UAT BOT ID') + 1).setValue(uatActiveBot);
          sheet.getRange(rowIndex, mainData.headers.indexOf('UAT JWT') + 1).setValue('JWT ' + uatToken);
          logLibraryUsage('Generate UAT Token', 'Pass');
          Logger.log("✅ UAT Token and BOT ID updated successfully");
        }
      } catch (error) {
        logLibraryUsage('Generate UAT Token', 'Fail', error.toString());
        Browser.msgBox('❌ UAT Login Error: ' + error.toString());
      }
    }

    // PROD Login
    if (prodEmail) {
      try {
        const prodToken = loginAndGetToken(prodEmail, password, ACL_Domain);
        if (prodToken) {
          const prodActiveBot = getActiveBot(prodToken, ACL_Domain);
          sheet.getRange(rowIndex, mainData.headers.indexOf('PROD BOT ID') + 1).setValue(prodActiveBot);
          sheet.getRange(rowIndex, mainData.headers.indexOf('PROD JWT') + 1).setValue('JWT ' + prodToken);
          logLibraryUsage('Generate PROD Token', 'Pass');
          Logger.log("✅ PROD Token and BOT ID updated successfully");
        }
      } catch (error) {
        logLibraryUsage('Generate PROD Token', 'Fail', error.toString());
        handleError(error);
//        Browser.msgBox('❌ PROD Login Error: ' + error.toString());
      }
    }

  } catch (error) {
    logLibraryUsage('Generate Token', 'Fail', error.toString());
    Logger.log('❌ Script failed in loginAndStoreToken: ' + error.toString());
    ss.toast('Error: ' + error.toString());
handleError(error);
  }
}


function loginAndGetToken(email, password, ACL_Domain) {
  try{
const endpoints = getApiEndpoints();
const loginUrl = ACL_Domain + getValidatedEndpoint(endpoints,'ACL Login'); 

Logger.log(loginUrl);
 // const loginUrl = ACL_Domain + '/api/users/login';
  Logger.log(`🔐 Logging in: ${email}`);

  const loginResponse = UrlFetchApp.fetch(loginUrl, {
    method: 'post',
    headers: { 'Content-Type': 'application/json' },
    payload: JSON.stringify({ email, password }),
    muteHttpExceptions: true
  });

  if (loginResponse.getResponseCode() === 200) {
    Logger.log(`✅ Login successful for: ${email}`);
    return JSON.parse(loginResponse.getContentText()).token;
  } else {
    throw new Error('Error logging in: ' + loginResponse.getResponseCode() + ' - ' + loginResponse.getContentText());
  }
  }
  catch(error)
  {
    handleError(error);
  }
}



function getActiveBot(token, ACL_Domain) {
try{
const endpoints = getApiEndpoints();
const url = ACL_Domain + getValidatedEndpoint(endpoints, 'ACL get Active Bot');
Logger.log(url)
  // const url = ACL_Domain + '/api/users/dashboard-init';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Authorization': 'JWT ' + token,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    Logger.log("✅ Retrieved activeBot from dashboard-init");
    return data.activeBot;
  } else {
    throw new Error('Error fetching activeBot: ' + response.getResponseCode() + ' - ' + response.getContentText());
  }
}catch(error)
{handleError(error);
}
}

