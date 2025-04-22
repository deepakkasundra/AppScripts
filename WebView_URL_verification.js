function fetch_WebviewURL_PROD() {
  try {
    fetchAndStoreWebViewURLs('PROD');
  } catch (e) {
    Logger.log('Error in fetch_WebviewURL_PROD: ' + e.message);
      Browser.msgBox(e.message, Browser.Buttons.OK);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error in fetch_WebviewURL_PROD: ' + e.message, 'Error', 5);
  }
}

function fetch_WebviewURL_UAT() {
  try {
    fetchAndStoreWebViewURLs('UAT');
  } catch (e) {
    Logger.log('Error in fetch_WebviewURL_UAT: ' + e.message);
    Browser.msgBox(e.message, Browser.Buttons.OK);

    SpreadsheetApp.getActiveSpreadsheet().toast('Error in fetch_WebviewURL_UAT: ' + e.message, 'Error', 5);
  }
}

function fetchAndStoreWebViewURLs(env) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('Main');

  try {
    // Fetch BOT ID and Domain Name from Main sheet
    var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];
    var rowIndex = 2;

    var botID = mainSheet.getRange(rowIndex, headersValues.indexOf(env + ' BOT ID') + 1).getValue();
    var dashboardDomain = mainSheet.getRange(rowIndex, headersValues.indexOf('Mono CM Ticketing form') + 1).getValue();

    Logger.log(botID);
    Logger.log(dashboardDomain);

    // Check if BOT ID or Domain name is blank
    if (!botID || !dashboardDomain) {
      throw new Error('BOT ID or Domain name is blank.');
    }

    // API details
    const url = `${dashboardDomain}/api/bots/${botID}/utils/urls`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `JWT eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI1ODg5OTUyYWIxOTMwNGE1ZTdlZWEyNWYiLCJuYW1lIjoiQ2hhdHRlck9uIiwiaWF0IjoxNTU4NTk3NTI5LCJleHAiOjE1NjM3ODE1Mjl9.Lgk6d4xyuXKvvJ3_vsbLJJ2ayXEod5o_rlms8FEDhsQ`
      }
    };

    // Fetch data from the API
    try {
      const response = UrlFetchApp.fetch(url, options);
      const data = JSON.parse(response.getContentText());

      // Log the fetched data
      Logger.log(data);

      // Retrieve the new format URL patterns from the sheet
      const newFormatURLs = getNewFormatURLsFromSheet(ss);
      Logger.log(newFormatURLs);

      // Initialize an array to hold all relevant URLs
      const urls = [];

      // Iterate through the data and capture URLs that match the patterns
data.forEach(item => {
  if (item.button && item.button.webView) {
    const title = item.button.webView.title;
    const url = item.button.webView.url;
    const name = item.name || 'N/A';  // Fallback to 'N/A' if name is not provided
    const pageTitle = item.button.webView.pageTitle || 'N/A';  // Fallback to 'N/A' if pageTitle is not provided

    // Determine the status and reason based on the URL format
    let status = 'NA';
    let reason = '';

    if (url.includes('Tickets') || url.includes('tickets') || url.includes('createTicket')) {
      status = 'Old Format';
      if (title in newFormatURLs) {
        const comparisonResult = compareURLs(newFormatURLs[title], url);
        if (comparisonResult.match) {
          status = 'New Format';
          reason = ''; // Clear reason as it's a match
        } else {
          reason = comparisonResult.reason; // Capture the mismatch reason
        }
      } else {
        reason = `Title "${title}" not found in NewFormatURLs`; // Reason if title is missing
      }
    }

    urls.push({
      particulars: title,
      url: url,
      module: item.id,
      status: status,
      reason: reason, // Add the reason field
      name: name,
      pageTitle: pageTitle
    });
  }
});


      // Check if "Configured_URL" sheet exists
      const WebViewSheet = env+'_WebView_URL_Verification';
      let sheet = ss.getSheetByName(WebViewSheet);

      if (sheet) {
        // Clear existing data
        sheet.clear();
      } else {
        // Create the sheet if it doesn't exist
        sheet = ss.insertSheet(WebViewSheet);
      }
ss.setActiveSheet(sheet);
      // Insert the headers and the filtered data into the sheet
      const headers = ['Particulars', 'Name', 'Page Title', 'URLS', 'Module', 'Status', 'Reason'];
      sheet.appendRow(headers);

      urls.forEach(item => {
        sheet.appendRow([item.particulars, item.name, item.pageTitle, item.url, item.module, item.status, item.reason]);
      });

      ss.toast('URLs have been successfully fetched and stored.', 'Task Completed', 5);
    } catch (e) {
      Logger.log('Error fetching data from the API: ' + e.message);
      Browser.msgBox(e.message, Browser.Buttons.OK);
      ss.toast('Error fetching data from the API: ' + e.message, 'API Fetch Error', 5);
    }

  } catch (e) {
    Logger.log('Error in fetchAndStoreWebViewURLs: ' + e.message);
    Browser.msgBox(e.message, Browser.Buttons.OK);
    ss.toast('Error in fetchAndStoreWebViewURLs: ' + e.message, 'Script Error', 5);
  }
}

// Function to retrieve newFormatURLs from a sheet
function getNewFormatURLsFromSheet(ss) {
  try {
    let sheet = ss.getSheetByName('NewFormatURLs');
    if (!sheet) {
      throw new Error('Sheet "NewFormatURLs" not found.');
    }

    const data = sheet.getDataRange().getValues();
    const newFormatURLs = {};

    for (let i = 1; i < data.length; i++) { // Start from 1 to skip headers
      const [title, url] = data[i];
      newFormatURLs[title] = url;
    }

    return newFormatURLs;
  } catch (e) {
    Logger.log('Error in getNewFormatURLsFromSheet: ' + e.message);
   Browser.msgBox(e.message, Browser.Buttons.OK);

    throw e;  // Rethrow the error to be caught in the main function
  }
}


function compareURLs(newFormatURL, retrievedURL) {
  try {
    Logger.log("Starting URL comparison");
    Logger.log("New Format URL: " + newFormatURL);
    Logger.log("Retrieved URL: " + retrievedURL);

    const newFormatBase = newFormatURL.split('?')[0];
    const retrievedBase = retrievedURL.split('?')[0];

    // Compare the base URLs
    if (newFormatBase !== retrievedBase) {
      Logger.log(`Base URLs do not match. Expected: ${newFormatBase}, Found: ${retrievedBase}`);
      return {
        match: false,
        reason: `Base URLs do not match. Expected: ${newFormatBase}, Found: ${retrievedBase}`,
        status: 'Old Format' // Status for mismatched base URLs
      };
    }

    Logger.log("Base URLs match");

    // Parse query parameters into objects for comparison
    let newFormatParams = parseQueryParams(newFormatURL);
    let retrievedParams = parseQueryParams(retrievedURL);

    Logger.log("Parsed New Format Params: " + JSON.stringify(newFormatParams));
    Logger.log("Parsed Retrieved Params: " + JSON.stringify(retrievedParams));

    // Ignore 'botid' and 'userid' in both sets of parameters
    const ignoredParams = ['botId', 'userId'];
    newFormatParams = filterIgnoredParams(newFormatParams, ignoredParams);
    retrievedParams = filterIgnoredParams(retrievedParams, ignoredParams);

    Logger.log("Filtered New Format Params: " + JSON.stringify(newFormatParams));
    Logger.log("Filtered Retrieved Params: " + JSON.stringify(retrievedParams));

    // Check for missing parameters in the retrieved URL
    const missingParams = Object.keys(newFormatParams).filter(
      key => !retrievedParams.hasOwnProperty(key)
    );

    if (missingParams.length > 0) {
      Logger.log("Missing parameters in retrieved URL: " + missingParams.join(', '));
      return {
        match: false,
        reason: `Missing parameters in retrieved URL: ${missingParams.join(', ')}`,
        status: 'Old Format'
      };
    }

    Logger.log("All required parameters are present in the retrieved URL");

    // Additional validation for "createTicket" URL
    if (retrievedBase.includes('createTicket')) {
      Logger.log("URL contains 'createTicket'");

      // Check if 'settingsName' is present instead of 'formName'
      if (retrievedParams.settingsName) {
        Logger.log("Invalid parameter 'settingsName' found in 'createTicket' URL");
        return {
          match: false,
          reason: `'createTicket' URL should not include 'settingsName' parameter. Found: ${Object.keys(retrievedParams).join(', ')}`,
          status: 'Old Format'
        };
      }

      // Check for 'formName' parameter
      if (retrievedParams.formName) {
        Logger.log("'formName' parameter found. URL is in new format.");
        return {
          match: true,
          reason: '',
          status: 'New Format'
        };
      }

      // Check for invalid extraFields with .value
      const paramMismatchReasons = [];
      Object.keys(retrievedParams).forEach(key => {
        if (key.startsWith('extraFields.') && key.includes('.value')) {
          paramMismatchReasons.push(`Invalid parameter found: ${key}. Should not include ".value".`);
        }
      });

      if (paramMismatchReasons.length > 0) {
        Logger.log("Invalid extraFields parameters found: " + paramMismatchReasons.join(' '));
        return {
          match: false,
          reason: paramMismatchReasons.join(' '),
          status: 'Old Format' // Status for invalid extraFields parameters
        };
      }

      // If no formName and there are extraFields, it's considered the New Format
      if (Object.keys(retrievedParams).some(key => key.startsWith('extraFields.'))) {
        Logger.log("'extraFields' parameters found. URL is in new format.");
        return {
          match: true,
          reason: '',
          status: 'New Format'
        };
      }
    }

    // Check for extra parameters in the retrieved URL
    const extraParams = Object.keys(retrievedParams).filter(
      key => !newFormatParams.hasOwnProperty(key)
    );

    if (extraParams.length > 0) {
      Logger.log("Extra parameters found in retrieved URL: " + extraParams.join(', '));
      return {
        match: false,
        reason: `Extra parameters found in retrieved URL: ${extraParams.join(', ')}. This is considered a new format.`,
        status: 'New Format'
      };
    }

    Logger.log("URLs match and are in the new format");
    return { match: true, reason: '', status: 'New Format' }; // URLs match and are in the new format
  } catch (e) {
    Logger.log('Error in compareURLs: ' + e.message);
    return { match: false, reason: 'Error during URL comparison: ' + e.message, status: 'Old Format' };
  }
}

// Utility function to parse query parameters
function parseQueryParams(url) {
  const queryString = url.split('?')[1] || '';
  const params = {};
  queryString.split('&').forEach(param => {
    const [key, value] = param.split('=');
    if (key) {
      params[key] = decodeURIComponent(value || '');
    }
  });
  Logger.log("Parsed query parameters: " + JSON.stringify(params));
  return params;
}

// Utility function to filter out ignored parameters
function filterIgnoredParams(params, ignoredParams) {
  const filteredParams = {};
  Object.keys(params).forEach(key => {
    if (!ignoredParams.includes(key)) {
      filteredParams[key] = params[key];
    } else {
      Logger.log(`Ignoring parameter: ${key}`);
    }
  });
  return filteredParams;
}
