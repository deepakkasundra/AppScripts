function Fetch_ticketing_Form_Mono_CM() {
try{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('Main');
  var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
  var headersValues = headersRange.getValues()[0];
  var rowIndex = 2; // Assuming values start from row 2


  // Find the index of the dynamic header name for Form IDs in the TicketingForm sheet
  var ticketingFormSheet = ss.getSheetByName('TicketingForm');
  var ticketingFormHeadersRange = ticketingFormSheet.getRange(1, 1, 1, ticketingFormSheet.getLastColumn());
  var ticketingFormHeadersValues = ticketingFormHeadersRange.getValues()[0];
  var formIdHeaderIndex = ticketingFormHeadersValues.indexOf('Form ID') + 1; // Adjust 'FormID' 
  var monoTicketingFormValueIndex = ticketingFormHeadersValues.indexOf('Mono Ticketing Form Value') + 1; // Adjust 'Mono Ticketing Form 


  // Get the value of the dynamic header name
  var formIdHeaderValue = ticketingFormSheet.getRange(1, formIdHeaderIndex).getValue();

  if (!formIdHeaderValue) {
    SpreadsheetApp.getActiveSpreadsheet().toast('FormID header not found in TicketingForm sheet. Script stopped.');
    return;
  }

  var PROD_BOTID = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD BOT ID') + 1).getValue();
  var mono_domain = mainSheet.getRange(rowIndex, headersValues.indexOf('Mono CM Ticketing form') + 1).getValue();
  var MonoprodJWT = mainSheet.getRange(rowIndex, headersValues.indexOf('Flow chatteron JWT') + 1).getValue();
  
  if (!PROD_BOTID || !mono_domain) {
    SpreadsheetApp.getActiveSpreadsheet().toast('PROD Bot ID or Domain Name blank. Script stopped.');
    return;
  }

const endpoints = getApiEndpoints();
// const url = mono_domain + getValidatedEndpoint(endpoints,'MONO CM ticketing form'); 
Domain Domain = Domain + '/<REDACTED_PATH>/' + Domain + Domain(Domain,'Domain Domain Domain Domain');
 
  Logger.log(formurl);
  var sheet = ss.getSheetByName('TicketingForm');
  var lastRow = sheet.getLastRow();

  // Get form IDs directly from the TicketingForm sheet based on the dynamic header index
  var formIdsRange = sheet.getRange(2, formIdHeaderIndex, lastRow - 1, 1);
  var formIds = formIdsRange.getValues();

  // Check if there are any non-empty cells for form IDs
  var nonEmptyCellFound = formIds.some(row => row[0]);


  if (!nonEmptyCellFound) {
    Browser.msgBox('Error', 'No Form ID Found in the sheet.', Browser.Buttons.OK);
    return;
  }

  // Process form IDs and retrieve data
  for (var i = 0; i < formIds.length; i++) {
    var formId = formIds[i][0];
    if (!formId) continue; // Skip if encountering an empty cell


    var url = formurl + formId;
Logger.log(url)
    var headers = {
      'Authorization':  MonoprodJWT,
      // 'Cookie': COOKIE
    };

    var options = {
      'method': 'get',
      'headers': headers
    };

    try {
      var response = UrlFetchApp.fetch(url, options);
      var jsonResponse = JSON.parse(response.getContentText());
      var downloadLink = jsonResponse.result;

      // Write download link in the "Mono Ticketing Form Value" column for the corresponding form ID
      sheet.getRange(i + 2, monoTicketingFormValueIndex).setValue(downloadLink || 'No link found');
      Logger.log('Download link retrieved successfully for form ID ' + formId);
    } catch (error) {
      Logger.log('Error fetching data for form ID ' + formId + ': ' + error);
//        SpreadsheetApp.getActiveSpreadsheet().toast('Error fetching data for form ID ' + formId, "Error", 5); // Display short error toast

    }
  }
}
catch(error)
{handleError(error);
}
}
