// Function to open the input dialog for pasting JSON
function openJsonToCsvDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Json_to_csv')
    .setWidth(700)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paste JSON Data');
}


function convertJsonToCsv(jsonText, sheetName, clearExisting) {
  return new Promise((resolve, reject) => {
    try {
      // Validate sheet name
      if (!sheetName) {
        return resolve({ success: false, message: 'Sheet name cannot be empty.' });
      }
      
      // const jsonData = JSON.parse(jsonText); // Parse the JSON
      
      // // Ensure the data array exists in the parsed JSON
      // if (!jsonData.data || !Array.isArray(jsonData.data)) {
      //   return resolve({ success: false, message: 'Invalid JSON format: "data" field is missing or not an array.' });
      // }
      // const dataArray = jsonData.data;
      

      let jsonData;
try {
  jsonData = JSON.parse(jsonText); // Parse the JSON

  // Check if the parsed data has a 'data' property that is an array
  if (Array.isArray(jsonData.data)) {
    // Use jsonData.data directly
    dataArray = jsonData.data;
  } else {
    return resolve({ success: false, message: 'Invalid JSON format: expected "data" to be an array.' });
  }
} catch (error) {
  return resolve({ success: false, message: 'Error parsing JSON: ' + error.message });
}


 //     console.log('Parsed JSON Data:', jsonData);
// console.log('Data Array:', dataArray);


      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = spreadsheet.getSheetByName(sheetName);

      if (sheet) {
        if (clearExisting) {
          // Clear the existing sheet if required
          sheet.clear();
        } else {
          // Return error if sheet already exists and clearExisting is not checked
          return resolve({ success: false, message: 'Sheet already exists. Uncheck the "Clear existing sheet" option to keep existing data.' });
        }
      } else {
        // Create a new sheet if it does not exist
        sheet = spreadsheet.insertSheet(sheetName);
      }

      // Flatten the JSON and prepare for CSV conversion
      const flatData = dataArray.map(record => flattenJson(record));
      
      // Extract headers from the flattened data
      const headers = getHeaders(flatData);

      // Create CSV data array with headers as the first row
      const csvData = [headers];

      // Add each row of data based on the headers
      flatData.forEach(item => {
        const row = headers.map(header => item[header] || ''); // Ensure every header has a value
        csvData.push(row);
      });

      // Load the data into the sheet
      sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

      // Return success message
      resolve({ success: true, message: 'Data successfully converted and loaded into sheet.' });
    } catch (error) {
      // Return error message
      reject({ success: false, message: 'Error parsing JSON or processing data: ' + error.message });
    }
  });
}


// Utility function to flatten a nested JSON object
function flattenJson(data, parentKey = '', result = {}) {
  for (const key in data) {
    const fullKey = parentKey ? `${parentKey}.${key}` : key;

    if (typeof data[key] === 'object' && data[key] !== null && !Array.isArray(data[key])) {
      // Recursively flatten nested objects
      flattenJson(data[key], fullKey, result);
    } else if (Array.isArray(data[key])) {
      // Flatten arrays
      data[key].forEach((item, index) => flattenJson(item, `${fullKey}[${index}]`, result));
    } else {
      // If primitive value, assign it to the result
      result[fullKey] = data[key];
    }
  }
  return result;
}

// Utility function to extract all unique headers from the flattened data
function getHeaders(flatData) {
  const headers = new Set();

  flatData.forEach(item => {
    Object.keys(item).forEach(key => headers.add(key));
  });

  return Array.from(headers); // Convert set to array
}

