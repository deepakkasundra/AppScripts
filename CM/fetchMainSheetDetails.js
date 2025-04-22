function getMainSheetData() {
  try {
//    Logger.log("🟢 Entered getMainSheetData()");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Main');
    if (!sheet) {
      throw new Error('Main sheet not found.');
    }

    const rowIndex = 2; // Assuming data is in the second row
//    Logger.log(`➡ Fetching headers from row 1`);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  //  Logger.log(`✅ Headers retrieved: ${JSON.stringify(headers)}`);

    const getValue = (headerName) => {
      const colIndex = headers.indexOf(headerName);
      if (colIndex === -1) {
        const errorMsg = `❌ Header "${headerName}" not found in Main sheet.`;
        Logger.log(errorMsg);
        throw new Error(errorMsg);
      }
      const value = sheet.getRange(rowIndex, colIndex + 1).getValue();
      // Logger.log(`📌 Value for "${headerName}": ${value}`);
      return value;
    };

    const prodEmail = getValue('PROD Email');
    const uatEmail = getValue('UAT Email');
    const password = getValue('Password');
    const prodBotId = getValue('PROD BOT ID');
    const prodJwt = getValue('PROD JWT');
    const uatBotId = getValue('UAT BOT ID');
    const uatJwt = getValue('UAT JWT');
//dashboard domain
    const domainname = getValue('Dashboard Domain Name');
    const ACL_Domain = getValue('ACL Domain Names');
    const flowchatteronDomain = getValue('Mono CM Ticketing form');
    const flowchatteronJWT = getValue('Flow chatteron JWT');
    //    Logger.log("✅ Successfully fetched all Main sheet values.");

    return {
      sheet,
      headers,
      rowIndex,
      getValue,
      prodEmail,
      uatEmail,
      password,
      prodBotId,
      prodJwt,
      uatBotId,
      uatJwt,
      domainname,
      ACL_Domain,
      flowchatteronDomain,
      flowchatteronJWT
    };

  } catch (error) {
    const msg = `❌ Error in getMainSheetData: ${error.message}`;
    Logger.log(msg);
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Error", 5);
    throw error;
  }
}

