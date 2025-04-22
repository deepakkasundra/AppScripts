/**
 * Entry points for UAT and PROD environments.
 */
function fetchDepartmentsFromUAT() {
  fetchAndStoreDepartments(false);
}

function fetchDepartmentsFromPROD() {
  fetchAndStoreDepartments(true);
}

/**
 * Fetches department details based on the selected environment and stores them in the "Department Details" sheet.
 * @param {Boolean} isProd - Pass true for PROD, false for UAT.
 */
function fetchAndStoreDepartments(isProd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  showProgressToast(ss, 'Initializing...');

  // Reuse data from getMainSheetData
  const mainData = getMainSheetData();
  var BOT_ID, jwt, Domain_name, sheetName;

  if (isProd) {
    BOT_ID = mainData.prodBotId;
    jwt = mainData.prodJwt;
    Domain_name = mainData.domainname;
    sheetName = "Department List_PROD";
  } else {
    BOT_ID = mainData.uatBotId;
    jwt = mainData.uatJwt;
    Domain_name = mainData.domainname;
    sheetName = "Department List_UAT";
  }

  
  // Validate required credentials.
  if (!BOT_ID || !jwt) {
    var envname = isProd ? "Production" : "UAT";
    Browser.msgBox(envname + " BOT ID or JWT Missing.", Browser.Buttons.OK);
    return;
  }

  showProgressToast(ss, 'Clearing Department Details sheet...');
  
  // Get or create the "Department Details" sheet and clear previous data.
  var deptSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  deptSheet.clear();

  // Construct the API URL dynamically.
  var url = Domain_name + '/bots/' + BOT_ID + '/@@@@@@@@@@@@@@@';
  var options = {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      "accept": "application/json, text/plain, */*",
      "authorization": jwt,
      "x-cm-dashboard-user": "true"
    }
  };

  try {
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

    var departments = data.data;
    if (departments.length === 0) {
      Logger.log("No department data found.");
      SpreadsheetApp.getUi().alert("No department data found.");
      return;
    }

    // Dynamically extract headers from the first object.
    var headers = Object.keys(departments[0]);
    deptSheet.appendRow(headers);


//   // Map the department data into rows.
// var rows = departments.map(function(dep) {
//   return headers.map(function(h) {
//     return dep[h] !== undefined && dep[h] !== null ? dep[h] : "";
//   });
// });

  // Map the TicketSchema data into rows.
var rows = departments.map(function(dep) {
  return headers.map(function(h) {
    var val = dep[h];
    if (val === null || val === undefined) return "";
    if (typeof val === "object") return JSON.stringify(val); // stringify arrays/objects
    return val;
  });
});

    deptSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    Logger.log("Department details fetched and stored successfully.");
    showProgressToast(ss, 'Department details fetched successfully.');
  } catch (error) {
    Logger.log("Error fetching department details: " + error.toString());
    SpreadsheetApp.getUi().alert("Error fetching department details: " + error.toString());
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
