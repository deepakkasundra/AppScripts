function runListFilesInFolderByUrl() {
    try {
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = spreadsheet.getSheetByName('Folder_Details');
        
        if (!sheet) {
            spreadsheet.toast('Sheet "Folder_Details" not found!', 'Error');
            return;
        }

        // Get headers from the first row
        var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
        var headersValues = headersRange.getValues()[0];
        var rowIndex = 2; // as value available at row 2
        var folderUrlColumnIndex = headersValues.indexOf('Folder Path (Drive URL)') + 1;
        if (folderUrlColumnIndex === 0) {
            spreadsheet.toast('Header "Folder Path (Drive URL)" not found!', 'Error');
            return;
        }

        var folderUrl = sheet.getRange(rowIndex, folderUrlColumnIndex).getValue();
        Logger.log(folderUrl);

        if (!folderUrl) {
            spreadsheet.toast('Folder URL is empty!', 'Error');
            return;
        }

        listFilesInFolderByUrl(folderUrl);
    } catch (e) {
        Logger.log('Error in runListFilesInFolderByUrl: ' + e.message);
        SpreadsheetApp.getActiveSpreadsheet().toast('An error occurred: ' + e.message, 'Error');
    }
}

function listFilesInFolderByUrl(folderUrl) {
    try {
        Logger.log('Folder URL: ' + folderUrl);

        // Extract the folder ID from the URL
        var folderId = folderUrl.match(/[-\w]{25,}/);
        if (!folderId) {
            Logger.log("Invalid folder URL: " + folderUrl);
            SpreadsheetApp.getActiveSpreadsheet().toast('Invalid folder URL', 'Error');
            return;
        }

        Logger.log("Folder ID: " + folderId[0]);

        var folder = DriveApp.getFolderById(folderId[0]);
        var files = folder.getFiles();
        
        if (!files.hasNext()) {
            Logger.log("Folder is empty or files are not accessible");
            SpreadsheetApp.getActiveSpreadsheet().toast('Folder is empty or files are not accessible', 'Error');
            return;
        }
        
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = spreadsheet.getSheetByName('Folder_Details_Data');

        if (!sheet) {
            // Create a new sheet named 'Folder_Details_Data' if it doesn't exist
            sheet = spreadsheet.insertSheet('Folder_Details_Data');
        } else {
            // Find the headers and clear content below dynamically
            var fileNameHeaderCell = sheet.createTextFinder("File Name").findNext();
            var urlHeaderCell = sheet.createTextFinder("URL").findNext();

            if (fileNameHeaderCell && urlHeaderCell) {
                var headerRow = fileNameHeaderCell.getRow();
                if (headerRow === urlHeaderCell.getRow()) {
                    // Clear the content below the headers
                     var numRowsToClear = sheet.getLastRow() - headerRow;
        if (numRowsToClear > 0) {
            sheet.getRange(headerRow + 1, 1, numRowsToClear, 2).clearContent();}
            }
            } else {
                // If headers are not found, create them in the top row (A1 and B1)
                sheet.getRange('A1').setValue("File Name");
                sheet.getRange('B1').setValue("URL");
            }
        }
        
        // Ensure headers are at the top row
        sheet.getRange('A1').setValue("File Name");
        sheet.getRange('B1').setValue("URL");
        
        var row = 2; // Start writing data from row 2, below the headers
        
        while (files.hasNext()) {
            var file = files.next();
            var fileName = file.getName();
            var fileUrl = file.getUrl();
            
            Logger.log("File Name: " + fileName + ", File URL: " + fileUrl);

            // Append file name and URL to the sheet starting from row 2
            sheet.getRange(row, 1).setValue(fileName);
            sheet.getRange(row, 2).setValue(fileUrl);
            row++;
        }
        
        Logger.log("Files listed in the sheet.");
        SpreadsheetApp.getActiveSpreadsheet().toast('Files successfully listed in the sheet.', 'Success');
    } 
    catch (e) {
        Logger.log('Error in listFilesInFolderByUrl: ' + e.message);
        SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}
