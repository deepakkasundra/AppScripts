function getAPIsheetDetails() {
  try {
    // Return the sheet details
    return {
      sheetId: '1pyvvrHVXXcZIiR_2CWsFDj5IkvYXNB5h86rDQjZWaJs',  // Replace with your actual sheet ID
      sheetName: 'API_EndPoints'  // Sheet name to monitor
    };
  }catch(e){
     handleError(e);
  }
  
  }

function getApiEndpoints() {
  try {
    // const sheetId = '1pyvvrHVXXcZIiR_2CWsFDj5IkvYXNB5h86rDQjZWaJs';
    // const sheetName = 'API_EndPoints';
    
 const { sheetId, sheetName } = getAPIsheetDetails();  // Get sheet details

    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in spreadsheet.`);

    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const particularsIndex = headers.indexOf('Particulars');
    const endpointIndex = headers.indexOf('EndPointName');

    if (particularsIndex === -1 || endpointIndex === -1) {
      throw new Error(`Required columns not found in "${sheetName}".`);
    }

    const endpoints = {};

    for (let i = 1; i < data.length; i++) {
      const key = data[i][particularsIndex];
      const value = data[i][endpointIndex];
      if (key) {
        endpoints[key] = value;
      }
    }

    Logger.log("✅ Endpoints fetched: " + JSON.stringify(endpoints));
    return endpoints;

  } catch (error) {
    const msg = `❌ Error in getApiEndpoints: ${error.message}`;
    Logger.log(msg);
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Error", 5);
     handleError(error);
    throw error;
  }
}

function getValidatedEndpoint(endpoints, key, context = '') {
 try{
  const endpointPath = endpoints[key];
  if (!endpointPath || typeof endpointPath !== 'string' || endpointPath.trim() === '') {
    const prefix = context ? `❌ ${context}:\n\n` : '';
    const errorMessage = 
      `${prefix}❌ Endpoint "${key}" is missing or not defined in the API_EndPoints sheet.\n\n` +
      `📩 Please connect with qa_manager@leena.ai to get the API Endpoint updated in the master sheet.`;

    // Show a pop-up toast in Google Sheets
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMessage, "Endpoint Error", 10); // 10 seconds toast

    // Throw the error to stop further execution
    throw new Error(errorMessage);
  }
  return endpointPath;
 }
 
 catch(error)
{
handleError(error);

} }



function validateAllEndpointFormats() {
  // const sheetId = '1pyvvrHVXXcZIiR_2CWsFDj5IkvYXNB5h86rDQjZWaJs';
  // const sheetName = 'API_EndPoints';
 const { sheetId, sheetName } = getAPIsheetDetails();  // Get sheet details


  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const particularsIndex = headers.indexOf('Particulars');
    const endpointIndex = headers.indexOf('EndPointName');

    // Add or find 'Validation Status' column
    let statusIndex = headers.indexOf('Validation Status');
    if (statusIndex === -1) {
      statusIndex = headers.length;
      sheet.getRange(1, statusIndex + 1).setValue('Validation Status');
    }

    let issuesFound = false;

    for (let i = 1; i < data.length; i++) {
      const key = data[i][particularsIndex];
      const endpointPath = data[i][endpointIndex];
      const validationResult = validateEndpoint(endpointPath, key);
      
      sheet.getRange(i + 1, statusIndex + 1).setValue(validationResult);
      
      if (validationResult !== '✅ Valid') {
        issuesFound = true;
      }
    }

    const toastMsg = issuesFound
      ? "⚠️ Some endpoints have invalid formatting. See 'Validation Status' column."
      : "✅ All endpoints are formatted correctly.";

    SpreadsheetApp.getActiveSpreadsheet().toast(toastMsg, "Validation Complete", 8);
    Logger.log(toastMsg);

  } catch (error) {
    const msg = `❌ Error during endpoint validation: ${error.message}`;
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Validation Error", 8);
    Logger.log(msg);
    handleError(error);
  }
}

function onEdit(e) {
  // const sheetId = '1pyvvrHVXXcZIiR_2CWsFDj5IkvYXNB5h86rDQjZWaJs';  // Replace with your actual sheet ID
  // const sheetName = 'API_EndPoints';  // Sheet name to monitor

try{
 const { sheetId, sheetName } = getAPIsheetDetails();  // Get sheet details

  const ss = e.source;
  const sheet = ss.getSheetByName(sheetName);
  const range = e.range;
  
  if (!sheet || range.getSheet().getName() !== sheetName) {
  //Logger.log(`Edit happened in sheet: ${range.getSheet().getName()}, but we're watching ${sheetName}`);
  return;
} const headers = sheet.getDataRange().getValues()[0];
  const endpointIndex = headers.indexOf('EndPointName');
  const statusIndex = headers.indexOf('Validation Status');

  if (endpointIndex === -1 || statusIndex === -1) return;  // Ensure we have the necessary columns

  // Check if the edit was made in the 'EndPointName' column (Column index from 1)
  if (range.getColumn() === endpointIndex + 1) {
    const row = range.getRow();
    const endpointPath = sheet.getRange(row, endpointIndex + 1).getValue();
    const particularKey = sheet.getRange(row, headers.indexOf('Particulars') + 1).getValue();

    const validationResult = validateEndpoint(endpointPath, particularKey);
    sheet.getRange(row, statusIndex + 1).setValue(validationResult);
  }
}
catch(error)
{
  handleError(error);
}
}


function validateEndpoint(endpointPath, particularKey) {
try{
  const errors = [];

  if (!particularKey || !endpointPath) {
    return '❌ Missing key or endpoint';
  }

  // Rule 1: Endpoint must start with a slash
  if (!endpointPath.startsWith('/')) {
    errors.push(`does not start with "/"`);
  }

  // Rule 2: Query parameter format check
  const query = endpointPath.split('?')[1];
  if (query) {
    const params = query.split('&');
    params.forEach(param => {
      const [k, v] = param.split('=');
      if (!v) {
        errors.push(`param "${k}" is missing a value`);
        return;
      }

      // Special check for 'filter' param
      if (k === 'filter') {
        const missingEncodingParts = [];
        if (!v.includes('%7B')) missingEncodingParts.push('%7B');
        if (!v.includes('%22')) missingEncodingParts.push('%22');
        if (!v.includes('%7D')) missingEncodingParts.push('%7D');

        if (missingEncodingParts.length > 0) {
          errors.push(`param "filter" is missing proper encoding (should contain ${missingEncodingParts.join(', ')}); value: "${v}"`);
        } else {
          try {
            const decoded = decodeURIComponent(v);
            JSON.parse(decoded);
          } catch (e) {
            const preview = (() => {
              try { return decodeURIComponent(v); } catch { return '[unable to decode]'; }
            })();
            errors.push(`param "filter" has invalid JSON after decoding → "${preview}"`);
          }
        }
      }
    });
  }

  // Rule 3: Invalid percent encoding check
  const badEncoding = endpointPath.match(/%[^0-9A-F]{2}|[^%]%[^0-9A-F]/i);
  if (badEncoding) {
    errors.push(`invalid percent encoding like "${badEncoding[0]}"`);
  }

  // Return the validation result
  if (errors.length > 0) {
    return `❌ Invalid: ${errors.join('; ')}`;
  } else {
    return '✅ Valid';
  }
}catch(error)
{
  handleError(error);
}
}

