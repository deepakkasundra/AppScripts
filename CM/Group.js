/**
 * Entry points for UAT and PROD environments.
 */
function fetchGroupsFromUAT() {
  fetchAndStoreGroups(false);
}

function fetchGroupsFromPROD() {
  fetchAndStoreGroups(true);
}

/**
 * Fetches Groups details based on the selected environment and stores them in the "Groups Details" sheet.
 * @param {Boolean} isProd - Pass true for PROD, false for UAT.
 */
function fetchAndStoreGroups(isProd) {

  try {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  showProgressToast(ss, 'Initializing...');

  // Reuse data from getMainSheetData
  const mainData = getMainSheetData();
  var BOT_ID, jwt, Domain_name, sheetName;

  if (isProd) {
    BOT_ID = mainData.prodBotId;
    jwt = mainData.prodJwt;
    Domain_name = mainData.domainname;
    sheetName = "Group_PROD";
  } else {
    BOT_ID = mainData.uatBotId;
    jwt = mainData.uatJwt;
    Domain_name = mainData.domainname;
    sheetName = "Group_UAT";
  }

  
  // Validate required credentials.
  if (!BOT_ID || !jwt) {
    var envname = isProd ? "Production" : "UAT";
    Browser.msgBox(envname + " BOT ID or JWT Missing.", Browser.Buttons.OK);
    return;
  }

  showProgressToast(ss, 'Clearing Groups Details sheet...');
  
  // Get or create the "Groups Details" sheet and clear previous data.
  var deptSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  deptSheet.clear();


  const endpoints = getApiEndpoints();
Domain Domain = Domain + '/<REDACTED_PATH>/' + Domain + Domain(Domain,'Domain');

  Logger.log(url);
  var options = {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      "accept": "application/json, text/plain, */*",
      "authorization": jwt,
      "x-cm-dashboard-user": "true"
    }
  };

    Logger.log("Fetching data from API: " + url);
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log("Response Code: " + responseCode);
    Logger.log("Raw Response: " + responseText);

    if (responseCode !== 200) {
      Logger.log("API request failed with status: " + responseCode);
      SpreadsheetApp.getUi().alert("API request failed with status: " + responseCode + "\n" +responseText);
      return;
    }

    var data = JSON.parse(responseText);
    if (!data || !data.data || !Array.isArray(data.data)) {
      Logger.log("Invalid API response format. Data object: " + JSON.stringify(data, null, 2));
      SpreadsheetApp.getUi().alert("Invalid API response format. Check logs for details.");
      return;
    }

    var Groupss = data.data;
    if (Groupss.length === 0) {
      Logger.log("No Groups data found.");
      SpreadsheetApp.getUi().alert("No Groups data found.");
      return;
    }

extractJSONAndAppendHeaders(Groupss, deptSheet);


//    deptSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    Logger.log("Groups details fetched and stored successfully.");
    showProgressToast(ss, 'Groups details fetched successfully.');
  } catch (error) {
    Logger.log("Error fetching Groups details: " + error.toString());
    SpreadsheetApp.getUi().alert("Error fetching Groups details: " + error.toString());
    handleError(error);

  }
  
}

/**
 * Displays a toast message for progress updates.
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet.
 * @param {string} message - The message to display.
 */
function showProgressToast(ss, message) {
  ss.toast(message, 'Progress', 3);
}



