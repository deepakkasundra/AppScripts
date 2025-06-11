function downloadLinksAndSaveToDrive_New_CM_UAT() {
  downloadLinksAndSaveToDrive("Link", "NewCM_UAT_Ticketing_forms_");
}

function downloadLinksAndSaveToDrive_New_CM_PROD() {
  downloadLinksAndSaveToDrive("Link", "NewCM_PROD_Ticketing_forms_");
}


function downloadLinksAndSaveToDrive_Mono_CM() {
  downloadLinksAndSaveToDrive("Mono Ticketing Form Value", "PROD_Ticketing_forms_");
}



function downloadLinksAndSaveToDrive(linkColumnName, folderPrefix) {
  try{
Logger.log("starting for " + folderPrefix);  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TicketingForm'); // Adjust the sheet name if needed
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
ss.setActiveSheet(sheet);
  var linkIndex = headers.indexOf(linkColumnName) + 1;
  var nameIndex = headers.indexOf("Name") + 1;
  var downloadStatusIndex = headers.indexOf("Download Status") + 1;
  var fileNameIndex = headers.indexOf("File Name") + 1;
  var urlIndex = headers.indexOf("URL") + 1;

  if (linkIndex < 1 || nameIndex < 1) {
    Logger.log("Column '" + linkColumnName + "' or 'Name' not found");
    return;
  }

  if (downloadStatusIndex < 1) {
    // If the "Download Status" column does not exist, create it at the end
    downloadStatusIndex = headers.length + 1;
    sheet.getRange(1, downloadStatusIndex).setValue("Download Status");
    headers.push("Download Status");
  }

  if (fileNameIndex < 1) {
    // If the "File Name" column does not exist, create it at the end
    fileNameIndex = headers.length + 1;
    sheet.getRange(1, fileNameIndex).setValue("File Name");
    headers.push("File Name");
  }

  if (urlIndex < 1) {
    // If the "URL" column does not exist, create it at the end
    urlIndex = headers.length + 1;
    sheet.getRange(1, urlIndex).setValue("URL");
    headers.push("URL");
  }

  var lastRow = sheet.getLastRow();
  var links = sheet.getRange(2, linkIndex, lastRow - 1).getValues();
  var names = sheet.getRange(2, nameIndex, lastRow - 1).getValues();

  // Get the parent folder of the current Google Sheet
  var fileId = ss.getId();
  var file = DriveApp.getFileById(fileId);
  var parentFolder = file.getParents().next();

  // Create a unique folder name with current date and time
  var now = new Date();
  var folderName = folderPrefix + Utilities.formatDate(now, Session.getScriptTimeZone(), 'ddMMyyyy_HH_mm');
  var folder = parentFolder.createFolder(folderName);

  showProgressToast(ss, "Starting download for " + (lastRow - 1) + " records...");

  links.forEach(function(link, index) {
    var url = link[0];
    var name = names[index][0];
    if (url && name) {
      try {
        var response = UrlFetchApp.fetch(url);
        var blob = response.getBlob();

        // Extract filename from URL
        var fileName = extractFileNameFromUrl(url);
        if (!fileName) {
          fileName = "Ticketing_Form" + (index + 1);
        }

        // Append name to the filename
        fileName = name + "_" + fileName;

        // Save the file to the folder
        var savedFile = folder.createFile(blob).setName(fileName);
        var savedFileUrl = savedFile.getUrl();
        Logger.log("Downloaded and saved: " + fileName);

        // Save the download status, file name, and URL back to the Google Sheet
        sheet.getRange(index + 2, downloadStatusIndex).setValue("Downloaded");
        sheet.getRange(index + 2, fileNameIndex).setValue(fileName);
        sheet.getRange(index + 2, urlIndex).setValue(savedFileUrl);
      } catch (error) {
        Logger.log("Error downloading file from URL " + url + ": " + error.message);
        sheet.getRange(index + 2, downloadStatusIndex).setValue("Error: " + error.message);
        SpreadsheetApp.getActiveSpreadsheet().toast('Error downloading file from URL ' + url + ': ' + error.message);
      }
    } else {
      sheet.getRange(index + 2, downloadStatusIndex).setValue("Missing URL or Name");
    }

    // Update progress message
    showProgressToast(ss, "Processed " + (index + 1) + " out of " + (lastRow - 1) + " records...");
  });

  showProgressToast(ss, "Download completed for " + (lastRow - 1) + " records.");
  }
  catch(error)
  {
    handleError(error);
  }
}

function extractFileNameFromUrl(url) {
  try{
  var match = url.match(/filename%3D%22([^%]+)%22/);
  if (match && match[1]) {
    return decodeURIComponent(match[1]);
  }
  return null;
}
  catch(error)
  {
    handleError(error);
  }
}

// function showProgressToast(ss, message) {
//   ss.toast(message, 'Progress', 5); // Display for 5 seconds
//   SpreadsheetApp.flush(); // Ensure the UI updates are pushed out immediately
// }

