function fetchEmailAutomationFromUAT() {
  fetchEmailData(false, 'automation');
}

function fetchEmailAutomationFromPROD() {
  fetchEmailData(true, 'automation');
}

function fetchEmailConfigurationFromUAT() {
  fetchEmailData(false, 'configuration');
}

function fetchEmailConfigurationFromPROD() {
  fetchEmailData(true, 'configuration');
}

/**
 * Fetches Email Automation or Configuration based on the selected environment and type.
 * @param {Boolean} isProd - true for PROD, false for UAT
 * @param {string} type - 'automation' or 'configuration'
 */
function fetchEmailData(isProd, type) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    showProgressToast(ss, 'Initializing...');

    // Get main sheet data using the getMainSheetData function
    const { 
      mainSheet, prodBotId, prodJwt, uatBotId, uatJwt, domainname 
    } = getMainSheetData();

    // Retrieve necessary values for the specified environment (PROD/UAT)
    var botID = isProd ? prodBotId : uatBotId;
    var jwt = isProd ? prodJwt : uatJwt;
    var dashboardDomain = domainname; // Common dashboard domain

    // Check if required values are missing
    if (!botID || !jwt || !dashboardDomain) {
      var envName = isProd ? "Production" : "UAT";
      SpreadsheetApp.getUi().alert(envName + " BOT ID, JWT, or Domain Name is missing.");
      return;
    }

    // Construct the endpoint URL based on the 'type' (automation/configuration)
    var endpoint = type === 'automation' 
      ? '/cm/email-automation/list?current=1&filter=%7B%22isDeactivated%22:false%7D&perPage=10000'
      : '/cm/email-configuration/list?current=1&filter=%7B%22isActive%22%3Atrue%7D&perPage=10000';

    var url = dashboardDomain + '/bots/' + botID + endpoint;

    // Determine the sheet name based on the 'type' and environment
    var sheetName = (type === 'automation' ? 'Email Rules_' : 'Email Config_') + (isProd ? 'PROD' : 'UAT');

    showProgressToast(ss, 'Fetching data from API...');
    var deptSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    deptSheet.clear();

    // Setup request options
    var options = {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        "accept": "application/json, text/plain, */*",
        "authorization": jwt,
        "x-cm-dashboard-user": "true"
      }
    };

    // Make the API request
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    // Handle API errors
    if (responseCode !== 200) {
      SpreadsheetApp.getUi().alert("API request failed with status: " + responseCode + "\n" + responseText);
      return;
    }

    var data = JSON.parse(responseText);
    if (!data || !data.data || !Array.isArray(data.data)) {
      SpreadsheetApp.getUi().alert("Invalid API response format. Check logs.");
      return;
    }

    var records = data.data;
    if (records.length === 0) {
      SpreadsheetApp.getUi().alert("No data found.");
      return;
    }

    // Extract headers from the first record and append to sheet
    var headers = Object.keys(records[0]);
    deptSheet.appendRow(headers);

    // Map the records into rows and insert into the sheet
    var rows = records.map(function(rec) {
      return headers.map(function(h) {
        var val = rec[h];
        return (val === null || val === undefined) ? "" : (typeof val === "object" ? JSON.stringify(val) : val);
      });
    });

    deptSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    showProgressToast(ss, type.charAt(0).toUpperCase() + type.slice(1) + ' fetched successfully.');

  } catch (e) {
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




// function fetchEmailAutomationFromUAT() {
//   fetchEmailData(false, 'automation');
// }

// function fetchEmailAutomationFromPROD() {
//   fetchEmailData(true, 'automation');
// }

// function fetchEmailConfigurationFromUAT() {
//   fetchEmailData(false, 'configuration');
// }

// function fetchEmailConfigurationFromPROD() {
//   fetchEmailData(true, 'configuration');
// }

// /**
//  * Fetches Email Automation or Configuration based on the selected environment and type.
//  * @param {Boolean} isProd - true for PROD, false for UAT
//  * @param {string} type - 'automation' or 'configuration'
//  */
// function fetchEmailData(isProd, type) {
//   try {
//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     showProgressToast(ss, 'Initializing...');

//     var mainSheet = ss.getSheetByName('Main');
//     var headersValues = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
//     var rowIndex = 2;

//     var BOT_ID = mainSheet.getRange(rowIndex, headersValues.indexOf(isProd ? 'PROD BOT ID' : 'UAT BOT ID') + 1).getValue();
//     var jwt = mainSheet.getRange(rowIndex, headersValues.indexOf(isProd ? 'PROD JWT' : 'UAT JWT') + 1).getValue();
//     var Domain_name = mainSheet.getRange(rowIndex, headersValues.indexOf('Dashboard Domain Name') + 1).getValue();

//     if (!BOT_ID || !jwt) {
//       var envName = isProd ? "Production" : "UAT";
//       SpreadsheetApp.getUi().alert(envName + " BOT ID or JWT Missing.");
//       return;
//     }

//     var endpoint = type === 'automation' 
//       ? '/cm/email-automation/list?current=1&filter=%7B%22isDeactivated%22:false%7D&perPage=10000'
//       : '/cm/email-configuration/list?current=1&filter=%7B%22isActive%22%3Atrue%7D&perPage=10000';

//     var url = Domain_name + '/bots/' + BOT_ID + endpoint;

//     var sheetName = (type === 'automation' ? 'Email Rules_' : 'Email Config_') + (isProd ? 'PROD' : 'UAT');

//     showProgressToast(ss, 'Fetching data from API...');
//     var deptSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
//     deptSheet.clear();

//     var options = {
//       method: "get",
//       muteHttpExceptions: true,
//       headers: {
//         "accept": "application/json, text/plain, */*",
//         "authorization": jwt,
//         "x-cm-dashboard-user": "true"
//       }
//     };

//     var response = UrlFetchApp.fetch(url, options);
//     var responseCode = response.getResponseCode();
//     var responseText = response.getContentText();

//     if (responseCode !== 200) {
//       SpreadsheetApp.getUi().alert("API request failed with status: " + responseCode + "\n" + responseText);
//       return;
//     }

//     var data = JSON.parse(responseText);
//     if (!data || !data.data || !Array.isArray(data.data)) {
//       SpreadsheetApp.getUi().alert("Invalid API response format. Check logs.");
//       return;
//     }

//     var records = data.data;
//     if (records.length === 0) {
//       SpreadsheetApp.getUi().alert("No data found.");
//       return;
//     }

//     var headers = Object.keys(records[0]);
//     deptSheet.appendRow(headers);

//     var rows = records.map(function(rec) {
//       return headers.map(function(h) {
//         var val = rec[h];
//         return (val === null || val === undefined) ? "" : (typeof val === "object" ? JSON.stringify(val) : val);
//       });
//     });

//     deptSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
//     showProgressToast(ss, type.charAt(0).toUpperCase() + type.slice(1) + ' fetched successfully.');

//   } catch (e) {
//     handleError(e);
//   }
// }

// /**
//  * Displays a toast message for progress updates.
//  * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet.
//  * @param {string} message - The message to display.
//  */
// function showProgressToast(ss, message) {
//   ss.toast(message, 'Progress', 3);
// }
