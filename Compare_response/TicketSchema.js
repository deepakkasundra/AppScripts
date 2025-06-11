/**
 * Entry points for UAT and PROD environments.
 */
function fetchTicketSchemaFromUAT() {
  fetchAndStoreTicketSchema(false);
}

function fetchTicketSchemaFromPROD() {
  fetchAndStoreTicketSchema(true);
}

/**
 * Fetches TicketSchema details based on the selected environment and stores them in the "TicketSchema Details" sheet.
 * @param {Boolean} isProd - Pass true for PROD, false for UAT.
 */
function fetchAndStoreTicketSchema(isProd) {
try
{

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    showProgressToast(ss, 'Initializing...');

    // Reuse data from getMainSheetData
    const mainData = getMainSheetData();
    var BOT_ID, jwt, Domain_name, sheetName;

    if (isProd) {
      BOT_ID = mainData.prodBotId;
      jwt = mainData.prodJwt;
      Domain_name = mainData.domainname;
      sheetName = "Ticketschema List_PROD";
    } else {
      BOT_ID = mainData.uatBotId;
      jwt = mainData.uatJwt;
      Domain_name = mainData.domainname;
      sheetName = "Ticketschema List_UAT";
    }

    // Validate required credentials.
    if (!BOT_ID || !jwt) {
      var envname = isProd ? "Production" : "UAT";
      Browser.msgBox(envname + " BOT ID or JWT Missing.", Browser.Buttons.OK);
      return;
    }

    showProgressToast(ss, 'Clearing TicketSchema Details sheet...');

    // Get or create the target sheet and clear existing data
    var deptSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    deptSheet.clear();

const endpoints = getApiEndpoints();
Domain Domain = Domain + '/<REDACTED_PATH>/' + Domain + Domain(Domain,'Domain Domain');

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

    var TicketSchema = data.data;
    if (TicketSchema.length === 0) {
      Logger.log("No TicketSchema data found.");
      SpreadsheetApp.getUi().alert("No TicketSchema data found.");
      return;
    }


// // Dynamically extract all unique headers from the TicketSchema array
// var headersSet = new Set();
// TicketSchema.forEach(dep => {
//   Object.keys(dep).forEach(key => headersSet.add(key));
// });
// var headers = Array.from(headersSet);

// // Add headers to the sheet
// deptSheet.appendRow(headers);


//   // Map the TicketSchema data into rows.
// var rows = TicketSchema.map(function(dep) {
//   return headers.map(function(h) {
//     var val = dep[h];
//     if (val === null || val === undefined) return "";
//     if (typeof val === "object") return JSON.stringify(val); // stringify arrays/objects
//     return val;
//   });
// });

extractJSONAndAppendHeaders(TicketSchema, deptSheet);


//    deptSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    Logger.log("TicketSchema details fetched and stored successfully.");
    showProgressToast(ss, 'TicketSchema details fetched successfully.');
  } catch (error) {
    Logger.log("Error fetching TicketSchema details: " + error.toString());
    SpreadsheetApp.getUi().alert("Error fetching TicketSchema details: " + error.toString());
  }
}
catch (e) {
handleError(e);
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

