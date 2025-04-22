function fetchFromUAT() {
  Fetch_ticketing_Form(false);
}

function fetchFromPROD() {
  Fetch_ticketing_Form(true);
}

function Fetch_ticketing_Form(isProd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  showProgressToast(ss, 'Initializing...');

  // Use centralized data getter
  const {prodBotId, uatBotId,prodJwt, uatJwt,domainname, } = getMainSheetData();

  // Select based on environment
  const envName = isProd ? 'Production' : 'UAT';
  const BOT_ID = isProd ? prodBotId : uatBotId;
  const jwt = isProd ? prodJwt : uatJwt;

  // Check if BOT ID or JWT is missing
  if (!BOT_ID || !jwt) {
    Browser.msgBox(envName + ' BOT ID or JWT Missing..', Browser.Buttons.OK);
    return;
  }

  const Domain_name = domainname;
  const url = `${Domain_name}/bots/${BOT_ID}/cm/ticket-form/list?perPage=1000&current=1&select=name%2Cdepartments.name&child=departments`;

  const headers = {
    'Authorization': jwt,
    'x-cm-dashboard-user': 'true',
  };

  const options = {
    'method': 'get',
    'headers': headers,
  };

  try {
    Logger.log(url);
    Logger.log(jwt);

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    showProgressToast(ss, 'Updating sheet...');
    updateSheet(data, Domain_name, BOT_ID, jwt);

    showProgressToast(ss, 'Data retrieved successfully from API.');
  } catch (error) {
    Logger.log('Error fetching data from API: ' + error);
    logLibraryUsage('Fetch Ticketing Form - ' + envName, 'Fail', error.toString());

    if (error.hasOwnProperty('response') && error.response) {
      const contentText = error.response.getContentText();
      const dialogMessage = 'An error occurred while fetching data from the API:\n\n' + contentText;
      SpreadsheetApp.getUi().alert('⚠️ Error fetching data from API', dialogMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ Unexpected error', 'An unexpected error occurred: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}




function updateSheet(data, Domain_name, BOT_ID, jwt) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the headers
  var ticketingsheet_headers = ['Form ID', 'Name', 'Departments', 'Link', 'Mono Ticketing Form Value']; // Headers for TicketingForm
  var uploadsheet_headers = ['Form ID', 'Name', 'Departments', 'Link', 'Mono Ticketing Form Value', 'Form File URL']; // Headers for Upload_TicketingForm
    
  // Update TicketingForm sheet
  var ticketingSheet = spreadsheet.getSheetByName('TicketingForm');
  if (!ticketingSheet) {
    // Create the sheet and add headers if it doesn't exist
    ticketingSheet = spreadsheet.insertSheet('TicketingForm');
    ticketingSheet.appendRow(ticketingsheet_headers);
  } else {
    // Clear all data except the headers, only if there's data below the headers
    spreadsheet.setActiveSheet(ticketingSheet);
    if (ticketingSheet.getLastRow() > 1) {
      ticketingSheet.getRange(2, 1, ticketingSheet.getLastRow() - 1, ticketingSheet.getLastColumn()).clearContent();
    }
  }

  // Update Upload_TicketingForm sheet
  var uploadSheet = spreadsheet.getSheetByName('Upload_TicketingForm');
  if (!uploadSheet) {
    // Create the sheet and add headers if it doesn't exist
    uploadSheet = spreadsheet.insertSheet('Upload_TicketingForm');
    uploadSheet.appendRow(uploadsheet_headers);
  } else {
    // Clear all data except the headers, only if there's data below the headers
    if (uploadSheet.getLastRow() > 1) {
      uploadSheet.getRange(2, 1, uploadSheet.getLastRow() - 1, uploadSheet.getLastColumn()).clearContent();
    }
  }

  // Write data to both sheets
  data.data.forEach(function(form, index) {
    showProgressToast(spreadsheet, 'Processing form ' + (index + 1) + ' of ' + data.data.length);

    // Prepare rowData for TicketingForm
    var rowDataTicketingForm = [];
    rowDataTicketingForm.push(form.formId || '');
    rowDataTicketingForm.push(form.name || '');

    // Extract department names if available
    var departmentNames = '';
    if (form.departments && form.departments.length > 0) {
      departmentNames = form.departments.map(function(dept) {
        return dept.name;
      }).join(', ');
    }
    rowDataTicketingForm.push(departmentNames);

    // Fetch link for each form ID and add it to the row
    var link = getLinkForFormId(form.formId, Domain_name, BOT_ID, jwt);
    rowDataTicketingForm.push(link);

    // Write the same rowData to the TicketingForm sheet
    ticketingSheet.appendRow(rowDataTicketingForm);

    // Prepare rowData for Upload_TicketingForm (Form URL should be blank)
    var rowDataUploadForm = rowDataTicketingForm.slice(); // Copy the same data
    rowDataUploadForm.push(''); // Leave Form URL blank

    // Write to Upload_TicketingForm sheet
    uploadSheet.appendRow(rowDataUploadForm);
  });
}


function getLinkForFormId(formId, Domain_name, BOT_ID, jwt) {
  var apiUrl = '' + Domain_name + '/bots/' + BOT_ID + '/cm/ticket-form/download-form?id=' + formId;
  Logger.log(apiUrl);
  var headers = {
    'Authorization': jwt,
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
    logLibraryUsage('Fetch tkt form getLinkForFormId' + formId, 'Fail', error.toString());
    return ''; // Return empty string if there's an error
  }
}


// function fetchFromUAT() {
 
//   Fetch_ticketing_Form(false);
// }

// function fetchFromPROD() {
//   Fetch_ticketing_Form(true);
// }

// function Fetch_ticketing_Form(isProd) {
 
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   showProgressToast(ss, 'Initializing...');

//   var mainSheet = ss.getSheetByName('Main');

//   var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
//   var headersValues = headersRange.getValues()[0];
//   var rowIndex = 2; // as value available at 2

//   var BOT_ID, jwt, Domain_name;

//   if (isProd) {
//     BOT_ID = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD BOT ID') + 1).getValue();
//     jwt = mainSheet.getRange(rowIndex, headersValues.indexOf('PROD JWT') + 1).getValue();
//   } else {
//     BOT_ID = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT BOT ID') + 1).getValue();
//     jwt = mainSheet.getRange(rowIndex, headersValues.indexOf('UAT JWT') + 1).getValue();
//   }

//  // Check if BOT_ID or JWT is missing
//     if (!BOT_ID || !jwt) {
//       if (isProd)
//       {
//         var envname = "Production";
//       }
//       else
//       {
//         var envname = "UAT";
//       }
//  Browser.msgBox(envname + " BOT ID or JWT Missing..", Browser.Buttons.OK);
// return;
//         }

//   Domain_name = mainSheet.getRange(rowIndex, headersValues.indexOf('Dashboard Domain Name') + 1).getValue();

//   showProgressToast(ss, 'Fetching data from API...');

//   var url = '' + Domain_name + '/bots/' + BOT_ID + '/cm/ticket-form/list?perPage=1000&current=1&select=name%2Cdepartments.name&child=departments';
//   var headers = {
//     'Authorization': jwt,
//     'x-cm-dashboard-user': 'true'
//   };

//   var options = {
//     'method': 'get',
//     'headers': headers
//   };
//   try {
//     Logger.log(url);
//     Logger.log(jwt);
//     var response = UrlFetchApp.fetch(url, options);
//     var json = response.getContentText();
//     var data = JSON.parse(json);

//     showProgressToast(ss, 'Updating sheet...');
//         updateSheet(data, Domain_name, BOT_ID, jwt);

// //    updateSheet(data, Domain_name, BOT_ID, jwt,isProd);

//     showProgressToast(ss, 'Data retrieved successfully from API.');
//   } catch (error) {
//     Logger.log('Error fetching data from API: ' + error);

//    logLibraryUsage('Fetch Tickting Form' + envname, 'Fail', error.toString());

//     if (error.hasOwnProperty('response') && error.response) {
//       var contentText = error.response.getContentText();
//       Logger.log('Full Error Message:', contentText);
//       var dialogTitle = '⚠️ Error fetching data from API';
//       var dialogMessage = 'An error occurred while fetching data from the API. Please see below for details:\n\n' + contentText;
//       SpreadsheetApp.getUi().alert(dialogTitle, dialogMessage, SpreadsheetApp.getUi().ButtonSet.OK);
//     } else {
//       // If it's not an API error, just display the error message directly
//       var dialogTitle = '⚠️ Unexpected error';
//       var dialogMessage = 'An unexpected error occurred: ' + error.toString();
//       SpreadsheetApp.getUi().alert(dialogTitle, dialogMessage, SpreadsheetApp.getUi().ButtonSet.OK);
//     }
  
//   }
// }
