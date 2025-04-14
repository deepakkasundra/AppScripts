function NLP_token_generate() 
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Find the NLP Token column dynamically
  var sheet = spreadsheet.getSheetByName('Main'); // Change 'Main' to your sheet name
  var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headersValues = headersRange.getValues()[0];
  var nlpTokenColumnIndex = headersValues.indexOf('NLP Token') + 1; // Adding 1 to convert from 0-based index to 1-based index
  var rowIndex = 2; // assuming value is available at row 2
  var nlp_user_name_value = sheet.getRange(rowIndex, headersValues.indexOf('NLP User Name') + 1).getValue();
  var nlp_user_Pass_value = sheet.getRange(rowIndex, headersValues.indexOf('NLP User Password') + 1).getValue();
  var NLP_URL_value = sheet.getRange(rowIndex, headersValues.indexOf('NLP Dashboard') + 1).getValue();
  var NLP_Cookie = '#########################';

  Logger.log(nlp_user_name_value + " & " + nlp_user_Pass_value);

  if (NLP_URL_value === "#N/A" || NLP_URL_value === "" || nlp_user_name_value === "" || nlp_user_Pass_value === "") {
    Logger.log("Either Domain name or NLP User Name or NLP User Password is blank. Further execution stopped.");
    SpreadsheetApp.getActiveSpreadsheet().toast('Either Domain name or NLP User Name OR Password missing', '⚠️ Further execution stopped.', 10);
    return;
  }

  if (nlpTokenColumnIndex === 0) {
    Logger.log("NLP Token column not found.");
    return;
  }

  try {
    var url = `${NLP_URL_value}/api/@@@@@@@@@@@@/`;
    var headers = {
      'Accept': '*/*',
      'Accept-Language': 'en-US,en;q=0.9',
      'Connection': 'keep-alive',
      'Content-Type': 'application/json',
      'Cookie': `${NLP_Cookie}/`, // Replace with actual cookie value if needed
      'Origin': `${NLP_URL_value}/`,
      'Referer': `${NLP_URL_value}/`,
      'Sec-Fetch-Dest': 'empty',
      'Sec-Fetch-Mode': 'cors',
      'Sec-Fetch-Site': 'same-origin',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
      'sec-ch-ua': '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
      'sec-ch-ua-mobile': '?0',
      'sec-ch-ua-platform': '"Windows"'
    };
    var payload = JSON.stringify({
      'username': nlp_user_name_value,
      'password': nlp_user_Pass_value
    });

    var options = {
      'method': 'post',
      'headers': headers,
      'payload': payload
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseData = JSON.parse(response.getContentText());

    // Extract the access token
    var accessToken = responseData.access;

    // Log the response for debugging
    Logger.log("Response Data: " + JSON.stringify(responseData));

 // Write access token to the NLP Token column
    var lastRow = sheet.getLastRow();
    var nlpTokenRange = sheet.getRange(2, nlpTokenColumnIndex, lastRow - 1, 1); // Assuming data starts from row 2
    var values = nlpTokenRange.getValues();
    var updated = false;
    
    Logger.log("Updating token in column index: " + nlpTokenColumnIndex + " and last row: " + lastRow);
  logLibraryUsage('Generate NLP Token', 'Pass');  // Log NLP Pass
    
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === "") {
        // If the cell is empty, update it with the new token
        sheet.getRange(i + 2, nlpTokenColumnIndex).setValue(accessToken);
        Logger.log("Access token stored in Main sheet at row " + (i + 2) + " and column " + nlpTokenColumnIndex);
        updated = true;
        break;
      } else if (values[i][0] !== accessToken) {
        // If the cell is not empty and does not match the new token, update it
        sheet.getRange(i + 2, nlpTokenColumnIndex).setValue(accessToken);
        Logger.log("Access token updated in Main sheet at row " + (i + 2) + " and column " + nlpTokenColumnIndex);
          SpreadsheetApp.getActiveSpreadsheet().toast('Access token appended to Main sheet at row ' + (i + 2) + ' and column ' + nlpTokenColumnIndex, 'Token generated Successfully', 10);
        updated = true;
        break;
      }
    }
    
    if (!updated) {
      // If no updates were made, append the new token at the end of the column
      sheet.getRange(lastRow + 1, nlpTokenColumnIndex).setValue(accessToken);
      Logger.log("Access token appended to Main sheet at row " + (lastRow + 1) + " and column " + nlpTokenColumnIndex);
      SpreadsheetApp.getActiveSpreadsheet().toast('Access token appended to Main sheet at row ' + (i + 2) + ' and column ' + nlpTokenColumnIndex, 'Token generated Successfully', 10);
    }
  } 
     catch (error) {
      // handleError(error);
  //  var errorMessage = 'Error Generating NLP Token\n';

    try {
      var responseText = error.message;
      var responseJson = JSON.parse(responseText.split('Truncated server response: ')[1] || '{}');
      var errorCode = responseText.match(/returned code (\d+)/);
      var detail = responseJson.detail || 'No detail available';

    //  errorMessage += 'Error Code: ' + (errorCode ? errorCode[1] : 'Unknown') + '\n';
     // errorMessage += 'Detail: ' + detail;
    } catch (jsonError) {
      // Fallback for cases where error parsing fails
handleError(error);
    Logger.log('Error loading file data: ' + error.message);  // Log error message

      // errorMessage += 'Detail: ' + error.message;
  logLibraryUsage('Generate NLP Token', 'Fail', error.toString());  // Log  failure

    }

//    Browser.msgBox(errorMessage, Browser.Buttons.OK);
    Logger.log("An error occurred: " + error);
  }
}
