function fetchUsersFromUAT() {
  fetchAndStoreUsers(false);
}

function fetchUsersFromPROD() {
  fetchAndStoreUsers(true);
}

function fetchAndStoreUsers(isProd) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    showProgressToast(ss, 'Initializing User Fetch...');

    const mainData = getMainSheetData();
    let BOT_ID, jwt, Domain_name, ACL_Domain,sheetName;
ACL_Domain = mainData.ACL_Domain;

Logger.log(mainData.ACL_Domain);

    if (isProd) {
      BOT_ID = mainData.prodBotId;
      jwt = mainData.prodJwt;
      Domain_name = mainData.domainname;
      sheetName = "ACL_Users_PROD";
    } else {
      BOT_ID = mainData.uatBotId;
      jwt = mainData.uatJwt;
      Domain_name = mainData.domainname;
      sheetName = "ACL_Users_UAT";
    }

    Logger.log("BOT_ID: " + BOT_ID);
    Logger.log("JWT (shortened): " + (jwt ? jwt.substring(0, 30) + "..." : "Not Provided"));
    Logger.log("Domain: " + Domain_name);
    Logger.log("Sheet Name: " + sheetName);

    if (!BOT_ID || !jwt) {
      const envname = isProd ? "Production" : "UAT";
      Browser.msgBox(envname + " BOT ID or JWT Missing.", Browser.Buttons.OK);
      return;
    }

    showProgressToast(ss, 'Clearing Users Sheet...');
    const userSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    userSheet.clear(); // Clear existing data

  const endpoints = getApiEndpoints();
    const baseUrl = ACL_Domain + getValidatedEndpoint(endpoints,'ACL User part1') + BOT_ID + getValidatedEndpoint(endpoints,'ACL User part2');

    Logger.log(baseUrl);
    let page = 1;
    const limit = 20;  // Fetch 10 records per page
   // let allUsers = [];
    let headers = []; // Store headers dynamically
    let totalRecordsProcessed = 0;

    // Fetch all pages
    while (true) {
      const url = `${baseUrl}?limit=${limit}&page=${page}`;
      Logger.log("Fetching URL: " + url);

      const options = {
        method: "get",
        muteHttpExceptions: true,
        headers: {
          "accept": "application/json, text/plain, */*",
          "authorization": jwt
        }
      };

      let response;
      try {
        response = UrlFetchApp.fetch(url, options);
      } catch (fetchError) {
        Logger.log("Fetch failed for page " + page + ": " + fetchError.toString());
        SpreadsheetApp.getUi().alert("Fetch error on page " + page + ": " + fetchError.toString());
        return;
      }

      const code = response.getResponseCode();
      const content = response.getContentText();
      Logger.log("HTTP Response Code: " + code);
      Logger.log("Response Content (first 500 chars):\n" + content.substring(0, 500));

      if (code !== 200) {
        SpreadsheetApp.getUi().alert(`API call failed on page ${page} with code ${code}\n${content}`);
        return;
      }

      let parsed;
      try {
        parsed = JSON.parse(content);
        Logger.log("Parsed keys: " + Object.keys(parsed).join(", "));
        Logger.log("Full response:\n" + JSON.stringify(parsed, null, 2));
      } catch (jsonError) {
        Logger.log("JSON Parsing Error: " + jsonError.toString());
        Logger.log("Raw Response:\n" + content);
        SpreadsheetApp.getUi().alert("JSON parsing error on page " + page + ". Check logs.");
        return;
      }

      if (!parsed || !parsed.users || !Array.isArray(parsed.users)) {
        Logger.log("Error: 'users' field is missing or invalid.");
        SpreadsheetApp.getUi().alert("Invalid response: 'users' array missing or malformed.");
        return;
      }

      // If no users found, stop the loop
      if (parsed.users.length === 0) {
        Logger.log("No users found on this page. Stopping.");
        break;
      }

      // Dynamically handle headers based on the API response
      const currentHeaders = Object.keys(parsed.users[0]);
      if (headers.length === 0) {
        headers = currentHeaders; // Initialize headers on the first page
        userSheet.appendRow(headers); // Write headers to the sheet first time
      } else if (headers.toString() !== currentHeaders.toString()) {
        headers = Array.from(new Set([...headers, ...currentHeaders])); // Merge new headers
        userSheet.getRange(1, 1, 1, headers.length).setValues([headers]); // Update headers in the sheet
      }

      // Adjust records to match headers
      const rows = parsed.users.map(user =>
        headers.map(header => user[header] !== undefined ? user[header] : '') // Fill missing fields with empty values
      );

      // Write data to the sheet
      const startRow = userSheet.getLastRow() + 1;
      userSheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);

      // Update progress logger
      totalRecordsProcessed += parsed.users.length;
      Logger.log(`Page ${page}: Processed ${parsed.users.length} records (Total: ${totalRecordsProcessed})`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Page ${page}: Processed ${parsed.users.length} records (Total: ${totalRecordsProcessed})`, 'Progress', 3);

      // Sleep for 5 seconds after processing a page
       Utilities.sleep(1000);  // Sleep for 5 seconds

      // Check if there is more data (pagination)
      const hasMoreData = parsed.users.length === limit; // If the page has records equal to the limit, assume there may be more pages
      if (!hasMoreData) {
        Logger.log(`No more pages. Data fetching completed.`);
        break; // Stop fetching if no more records
      }

      page++; // Move to the next page
    }

    showProgressToast(ss, `Fetched ${totalRecordsProcessed} users successfully.`);
    Logger.log("User data written to sheet: " + sheetName);

  } catch (error) {
    Logger.log("Unexpected Error: " + error.toString());
    SpreadsheetApp.getUi().alert("Unexpected error: " + error.toString());
    handleError(error);
  }
}

