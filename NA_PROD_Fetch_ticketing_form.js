
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('Main');

    var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
		var headersValues = headersRange.getValues()[0];
		var rowIndex = 2; // as value availabe at 2
	var PROD_BOTID = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD BOT ID') + 1).getValue();
    var prodJwt = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD JWT') + 1).getValue();
     
//    var PROD_BOTID = mainSheet.getRange('D2').getValue();
   // var prodJwt = mainSheet.getRange('E2').getValue();
//    var Domain_name = 'https://case-management-api.leena.ai'
// var Domain_name = mainSheet.getRange('H2').getValue();
var Domain_name = mainSheet.getRange(rowIndex, headersValues.indexOf('Dashboard Domain Name') + 1).getValue();

function Fetch_ticketing_Form_From_PROD() {
// Logger.log("E2 should be" + prodJwt)
Logger.log(PROD_BOTID);
  var url = ''+Domain_name+'/bots/'+PROD_BOTID+'/cm/ticket-form/list?perPage=1000&current=1&select=name%2Cdepartments.name&child=departments';
  var headers = {
    'Authorization': prodJwt,
    'x-cm-dashboard-user': 'true'
  };

  var options = {
    'method': 'get',
    'headers': headers
  };
  try {

Logger.log(url);
Logger.log(prodJwt)
    var response = UrlFetchApp.fetch(url, options);
    var json = response.getContentText();
    var data = JSON.parse(json);

    updateSheet_PROD(data);
    Logger.log('Data retrieved successfully from API.');
  }
  catch (error) {
  Logger.log('Error fetching data from API: ' + error);
  
  if (error.hasOwnProperty('response') && error.response) {
    var contentText = error.response.getContentText();
    Logger.log('Full Error Message:', contentText);
    var ui = SpreadsheetApp.getUi();
    var dialogTitle = '⚠️ Error fetching data from API';
    var dialogMessage = 'An error occurred while fetching data from the API. Please see below for details:\n\n' + contentText;
    ui.alert(dialogTitle, dialogMessage, ui.ButtonSet.OK);
  } else {
    // If it's not an API error, just display the error message directly
    var ui = SpreadsheetApp.getUi();
    var dialogTitle = '⚠️ Unexpected error';
    var dialogMessage = 'An unexpected error occurred: ' + error.toString();
    ui.alert(dialogTitle, dialogMessage, ui.ButtonSet.OK);
  }
}


}
function updateSheet_PROD(data) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('TicketingForm');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('TicketingForm');
  } else {
    // Clear existing data
sheet.clear();
  }

  // Write headers
  var headers = ['Form ID', 'Name', 'Departments', 'Link', 'Mono Ticketing Form Value']; // Added 'Link' header
  sheet.appendRow(headers);

  // Write data
  data.data.forEach(function(form) {
    var rowData = [];
    rowData.push(form.formId || '');
    rowData.push(form.name || '');

    // Extract department names if available
    var departmentNames = '';
    if (form.departments && form.departments.length > 0) {
      departmentNames = form.departments.map(function(dept) {
        return dept.name;
      }).join(', ');
    }
    rowData.push(departmentNames);

    // Fetch link for each form ID and add it to the row
    var link = getLinkForFormId_PROD(form.formId);
    rowData.push(link);

    sheet.appendRow(rowData);
  });
}

function getLinkForFormId_PROD(formId) {
  var apiUrl = ''+Domain_name+'/bots/'+PROD_BOTID+'/cm/ticket-form/download-form?id=' + formId;
  Logger.log(apiUrl)
  var headers = {
    'Authorization': prodJwt,
    'x-cm-dashboard-user': 'true'
  };

  var options = {
    'method': 'get',
    'headers': headers
  };

  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var json = response.getContentText();
    var data = JSON.parse(json);
    return data.url; // Assuming the API returns a 'url' property for the link
  } catch (error) {
    Logger.log('Error fetching link for form ID ' + formId + ': ' + error);
    return ''; // Return empty string if there's an error
  }
}
