function replaceMultipleValuesInXLSX() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Folder_Details');
  var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headersValues = headersRange.getValues()[0];
  var rowIndex = 2; // as value available at 2
  var folderUrl = sheet.getRange(rowIndex, headersValues.indexOf('Folder Path (Drive URL)') + 1).getValue();
  var exactMatch = sheet.getRange(rowIndex, headersValues.indexOf('Exact Match') + 1).getValue();

  Logger.log(folderUrl);
  Logger.log(exactMatch);

  Logger.log('Fetching folder ID from URL...');
  const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
  if (!folderIdMatch) {
    Logger.log('Invalid Folder URL');
    return;
  }
  const folderId = folderIdMatch[0];

  Logger.log('Accessing folder with ID: ' + folderId);
  const folder = DriveApp.getFolderById(folderId);

  // Create a new folder with current timestamp
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy_HH-mm-ss");
  const newFolderName = 'Processed_files_' + formattedDate;
  const newFolder = folder.createFolder(newFolderName);
  Logger.log('Created new folder: ' + newFolderName);

  const files = folder.getFiles();
  let totalFiles = 0;

  while (files.hasNext()) {
    files.next();
    totalFiles++;
  }

  if (totalFiles === 0) {
    Logger.log('No files found in the folder.');
    return;
  }

  const fileIterator = folder.getFiles();
  let processedFiles = 0;

  const headers = sheet.getRange('A1:Z1').getValues()[0];
  const oldValueIndex = headers.indexOf('Old Value');
  const newValueIndex = headers.indexOf('New Value');

  if (oldValueIndex === -1 || newValueIndex === -1) {
    Logger.log('Old Value or New Value columns not found');
    return;
  }

  Logger.log('Old Value column index: ' + oldValueIndex);
  Logger.log('New Value column index: ' + newValueIndex);

  // Find the last non-empty row
  const oldValuesRange = sheet.getRange(2, oldValueIndex + 1, sheet.getLastRow() - 1).getValues();
  const newValuesRange = sheet.getRange(2, newValueIndex + 1, sheet.getLastRow() - 1).getValues();

  let lastOldValueRow = 0;
  let lastNewValueRow = 0;

  // Find the last non-empty row for Old Values
  for (let i = oldValuesRange.length - 1; i >= 0; i--) {
    if (oldValuesRange[i][0]) {
      lastOldValueRow = i + 2; // Account for header row and zero-indexing
      break;
    }
  }

  // Find the last non-empty row for New Values
  for (let i = newValuesRange.length - 1; i >= 0; i--) {
    if (newValuesRange[i][0]) {
      lastNewValueRow = i + 2; // Account for header row and zero-indexing
      break;
    }
  }

  // Use the smaller of the two rows
  const lastNonEmptyRow = Math.min(lastOldValueRow, lastNewValueRow);

  const replacements = [];
  for (let i = 2; i <= lastNonEmptyRow; i++) {
    const oldValue = sheet.getRange(i, oldValueIndex + 1).getValue();
    const newValue = sheet.getRange(i, newValueIndex + 1).getValue();
    if (oldValue && newValue) {
      replacements.push({ oldValue, newValue });
    }
  }

  // Get or create Replacement_status sheet
  let statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Replacement_status');
  if (!statusSheet) {
    statusSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Replacement_status');
    statusSheet.getRange('A1:E1').setValues([['File name', 'File URL', 'No of updation done', 'Status', 'Updated Positions']]);
  } else {
    statusSheet.clear(); // Clear existing data but keep headers
    statusSheet.getRange('A1:E1').setValues([['File name', 'File URL', 'No of updation done', 'Status', 'Updated Positions']]);
  }

  let row = 2;

  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    processedFiles++;
    const fileName = file.getName();
    const fileUrl = file.getUrl();
    let updateCount = 0;
    let updatedPositions = [];

    Logger.log(`Processing file (${processedFiles}/${totalFiles}): ` + fileName + ' (URL: ' + fileUrl + ')');
    SpreadsheetApp.getActiveSpreadsheet().toast(`Processing file ${processedFiles} of ${totalFiles}: ${fileName}`);

    try {
      const fileBlob = file.getBlob();
      Logger.log('Converting XLSX file to Google Sheets...');

      const resource = {
        title: fileName,
        mimeType: MimeType.GOOGLE_SHEETS
      };
      const newFile = Drive.Files.insert(resource, fileBlob, { convert: true });
      const googleSheetId = newFile.id;
      const googleSheetUrl = `https://docs.google.com/spreadsheets/d/${googleSheetId}/edit`;
      Logger.log('Converted file to Google Sheets with URL: ' + googleSheetUrl);

      const googleSheet = SpreadsheetApp.openById(googleSheetId);
      const sheets = googleSheet.getSheets();

      for (const sheet of sheets) {
        Logger.log('Processing sheet: ' + sheet.getName());
        const data = sheet.getDataRange().getValues();
        const headers = data[0];

        for (let i = 1; i < data.length; i++) {
          for (let j = 0; j < data[i].length; j++) {
            for (const replacement of replacements) {
              if ((exactMatch && data[i][j] === replacement.oldValue) || 
                  (!exactMatch && typeof data[i][j] === 'string' && data[i][j].includes(replacement.oldValue))) {
                sheet.getRange(i + 1, j + 1).setValue(data[i][j].replace(replacement.oldValue, replacement.newValue));
                updateCount++;
                const columnName = headers[j] || 'Column ' + (j + 1);
                const cellPosition = `${String.fromCharCode(65 + j)}${i + 1}`;
                if (!updatedPositions.includes(cellPosition)) {
                  updatedPositions.push(cellPosition);
                }
                Logger.log(`Updated cell at ${cellPosition} (Row ${i + 1}, Column ${columnName})`);
              }
            }
          }
        }
      }

      Logger.log('Waiting for 10 seconds to ensure changes are applied...');
      Utilities.sleep(10000); // 10 seconds delay

      // Verification step
      Logger.log('Verifying replacements...');
      let verificationErrors = [];
      for (const sheet of sheets) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          for (let j = 0; j < data[i].length; j++) {
            for (const replacement of replacements) {
              if (data[i][j] === replacement.oldValue) {
                const cellPosition = `${String.fromCharCode(65 + j)}${i + 1}`;
                verificationErrors.push(`Expected replacement at ${cellPosition} but found ${replacement.oldValue}`);
              }
            }
          }
        }
      }

      if (verificationErrors.length > 0) {
        Logger.log('Verification failed: ' + verificationErrors.join('; '));
        throw new Error('Replacement verification failed.');
      } else {
        Logger.log('All replacements verified successfully.');
      }

      Logger.log('Saving changes and converting Google Sheets back to XLSX...');

      const exportUrl = `https://docs.google.com/spreadsheets/d/${googleSheetId}/export?format=xlsx`;
      const response = UrlFetchApp.fetch(exportUrl, {
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
      });
      const xlsxBlob = response.getBlob().setName(fileName);

      const newXlsxFile = newFolder.createFile(xlsxBlob);
      const newXlsxFileUrl = newXlsxFile.getUrl();
      Logger.log('Created new XLSX file with URL: ' + newXlsxFileUrl);

      statusSheet.getRange(row, 1, 1, 5).setValues([
        [fileName, newXlsxFileUrl, updateCount, 'Success', updatedPositions.join(', ')]
      ]);
      
      // Flush the changes to ensure the updates are applied immediately
      SpreadsheetApp.flush();

    } catch (error) {
      Logger.log('Error processing file: ' + error.message);
      statusSheet.getRange(row, 1, 1, 5).setValues([
        [fileName, fileUrl, updateCount, 'Failed', error.message]
      ]);
    }
    row++;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Processing complete. ${processedFiles} files processed.`);
  Logger.log(`Processing complete. ${processedFiles} files processed.`);
}

