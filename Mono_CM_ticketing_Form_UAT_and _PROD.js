function MONO_fetchTicketingFormDataAndUrls_PROD() {
  MONO_fetchFormDataAndUrls('PROD');
}

function MONO_fetchTicketingFormDataAndUrls_UAT() {
  MONO_fetchFormDataAndUrls('UAT');
}

function MONO_fetchFormDataAndUrls(environment) {
  // var EnvJWT = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2MjlkYWRmYmEwZjY2YTRjMWExN2YxYjQiLCJuYW1lIjoiRGVlcGFrIFAgS2FzdW5kcmEiLCJpYXQiOjE3MjM2MDk1OTIsImV4cCI6MTcyNDgxOTE5MiwiaXNzIjoiY29yZSJ9.JeHsf15aWVgMjRsNK_01vxq3-ceTlVs5jpGJBCLTscw';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('Main');
  var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
  var headersValues = headersRange.getValues()[0];
  var rowIndex = 2; // Assuming values start from row 2

  var BOTID = mainSheet.getRange(rowIndex, headersValues.indexOf(environment + ' BOT ID') + 1).getValue();
  var mono_domain = mainSheet.getRange(rowIndex, headersValues.indexOf('Mono CM Ticketing form') + 1).getValue();
  var EnvJWT = mainSheet.getRange(rowIndex, headersValues.indexOf('Flow chatteron JWT') + 1).getValue();

  if (!BOTID || !mono_domain) {
    SpreadsheetApp.getActiveSpreadsheet().toast(environment + ' Bot ID or Domain Name blank. Script stopped.');
    return;
  }

  // Find the index of the dynamic header name for Form IDs in the TicketingForm sheet
  var ticketingFormSheet = ss.getSheetByName('TicketingForm');
  var ticketingFormHeadersRange = ticketingFormSheet.getRange(1, 1, 1, ticketingFormSheet.getLastColumn());
  var ticketingFormHeadersValues = ticketingFormHeadersRange.getValues()[0];
  var formIdHeaderIndex = ticketingFormHeadersValues.indexOf('Form ID') + 1; // Adjust 'Form ID' to your actual header name
  var monoTicketingFormValueIndex = ticketingFormHeadersValues.indexOf('Mono Ticketing Form Value') + 1; // Adjust 'Mono Ticketing Form Value' to your actual header name

  // Get the value of the dynamic header name
  var formIdHeaderValue = ticketingFormSheet.getRange(1, formIdHeaderIndex).getValue();

  if (!formIdHeaderValue) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Form ID header not found in TicketingForm sheet. Script stopped.');
    return;
  }

  var apiUrl = mono_domain + '/api/v1.0/bots/' + BOTID + '/tickets/settings/form-configs';
  Logger.log(apiUrl);

  // Authorization header
  var headers = {
    'Authorization': EnvJWT,
    'Accept': 'application/json'
  };

  try {
    // Fetch data from API
    var response = UrlFetchApp.fetch(apiUrl, { headers: headers });
    var content = response.getContentText();
    Logger.log('API Response: ' + content);

    // Parse response JSON
    var data = JSON.parse(content).result;

    // Access TicketingForm sheet
    var sheet = ticketingFormSheet;

    // Clear existing data in the sheet
    sheet.clear();

    // Write headers
    var headers = ['Form ID', 'Name', 'Departments', 'Link', 'Mono Ticketing Form Value']; // Added 'Link' header
    sheet.appendRow(headers);

    // Write new data to the sheet
    var rowData = [];

    data.forEach(function (form) {
      rowData.push([
        form._id, // Form ID
        form.name, // Name
        form.instanceName, // Departments - You may need to fetch this information from your API if available
        '', // Link - You may need to construct this based on form._id or other information
        '' // Mono Ticketing Form Value - This is currently empty, you may need to fill it based on your requirements
      ]);
    });

    // Set Form IDs and Names
    sheet.getRange(2, formIdHeaderIndex, rowData.length, rowData[0].length).setValues(rowData);

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

      var url = mono_domain + '/api/v1.0/bots/' + BOTID + '/tickets/settings/form-configs-file/' + formId;

      var headers = {
        'Authorization': EnvJWT,
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
 Browser.msgBox('API Error', error, Browser.Buttons.OK);
        Logger.log('Error fetching data for form ID ' + formId + ': ' + error);
      }
    }
  } catch (error) {
 Browser.msgBox('API Error', error, Browser.Buttons.OK);
    Logger.log('Error fetching data: ' + error);
  }
}

