function Conditional_replaceMultipleValuesInXLSX() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Folder_Details');
     var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];
    var rowIndex = 2; // as value availabe at 2
    var folderUrl = sheet.getRange(rowIndex, headersValues.indexOf('Folder Path (Drive URL)') + 1).getValue();

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

  // Dynamically find column indices for Key, Header, and New Value
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyHeaderIndex = headers.indexOf('Key');
  const headerNameIndex = headers.indexOf('Header');
  const newValueIndex = headers.indexOf('Set Value');

  if (keyHeaderIndex === -1 || headerNameIndex === -1 || newValueIndex === -1) {
    Logger.log('Key, Header, or Set Value column not found.');
    return;
  }

  // Read Key, Header, Set Value replacements from Folder_Details sheet
  const replacements = [];
  const replacementDataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < replacementDataRange.length; i++) {
    const key = replacementDataRange[i][keyHeaderIndex];
    const header = replacementDataRange[i][headerNameIndex];
    let newValue = replacementDataRange[i][newValueIndex];

    // Always convert newValue to string
    if (newValue !== undefined) {
      newValue = newValue.toString();
    }

    Logger.log(key + header + newValue)

    if (key && header && newValue !== undefined) {
      replacements.push({ key, header, newValue });
    }
  }

  // Get or create Replacement_status sheet
  let statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Replacement_status');
  if (!statusSheet) {
    statusSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Replacement_status');
    statusSheet.getRange('A1:E1').setValues([['File name', 'File URL', 'No of updates done', 'Status', 'Updated Positions']]);
  } else {
    statusSheet.clear(); // Clear existing data but keep headers
    statusSheet.getRange('A1:E1').setValues([['File name', 'File URL', 'No of updates done', 'Status', 'Updated Positions']]);
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
        const fileHeaders = data[0]; // Ensure fileHeaders is set here

        for (let i = 1; i < data.length; i++) {
          const rowKey = data[i][fileHeaders.indexOf('key')]; // Find the key in the current row

          for (let j = 0; j < fileHeaders.length; j++) {
            for (const replacement of replacements) {
              if (rowKey === replacement.key && fileHeaders[j] === replacement.header) {
                // Update the value based on key-header match
                sheet.getRange(i + 1, j + 1).setValue(replacement.newValue);
                updateCount++;
                const cellPosition = `${String.fromCharCode(65 + j)}${i + 1}`;
                if (!updatedPositions.includes(cellPosition)) {
                  updatedPositions.push(cellPosition);
                }
                Logger.log(`Updated cell at ${cellPosition} (Row ${i + 1}, Column ${replacement.header})`);
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
        const fileHeaders = data[0]; // Ensure fileHeaders is set here

        for (let i = 1; i < data.length; i++) {
          for (let j = 0; j < data[i].length; j++) {
            const header = fileHeaders[j];
            const key = data[i][fileHeaders.indexOf('key')];
            for (const replacement of replacements) {
              if (key === replacement.key && header === replacement.header) {
                const cellValue = data[i][j].toString(); // Convert cell value to string
                if (cellValue !== replacement.newValue) {
                  const cellPosition = `${String.fromCharCode(65 + j)}${i + 1}`;
                  verificationErrors.push(`Expected replacement at ${cellPosition} for key ${key} and header ${header} but found ${cellValue}`);
                }
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
      Logger.log('Created new XLSX file: ' + newXlsxFile.getName() + ' (URL: ' + newXlsxFileUrl + ')');

     // Only add row if important fields are not empty
  if (fileName && newXlsxFileUrl) {
    statusSheet.getRange(`A${row}`).setValue(fileName);
    statusSheet.getRange(`B${row}`).setValue(newXlsxFileUrl);
    statusSheet.getRange(`C${row}`).setValue(updateCount);
    statusSheet.getRange(`D${row}`).setValue('Updated Successfully');
    statusSheet.getRange(`E${row}`).setValue(updatedPositions.join(', '));
    row++;
  }

      // Delete the temporary Google Sheets file
      Logger.log('Deleting temporary Google Sheets file...');
      DriveApp.getFileById(googleSheetId).setTrashed(true);
      Logger.log('Deleted temporary Google Sheets file with URL: ' + googleSheetUrl);
    SpreadsheetApp.flush(); // Force update to the sheet
    Logger.log('Updates flushed to the sheet.');


    } catch (e) {
      Logger.log('Error processing file ' + fileName + ': ' + e.message);
      statusSheet.getRange(`A${row}`).setValue(fileName);
      statusSheet.getRange(`B${row}`).setValue(fileUrl);
      statusSheet.getRange(`C${row}`).setValue(updateCount);
      statusSheet.getRange(`D${row}`).setValue('Error: ' + e.message);
      statusSheet.getRange(`E${row}`).setValue('N/A');
      row++;
    SpreadsheetApp.flush(); // Force update to the sheet
    Logger.log('Updates flushed to the sheet.');

    }
  }

  Logger.log('File processing complete.');
    SpreadsheetApp.getActiveSpreadsheet().toast('File processing complete. Please refer "Replacement_status" ', 'Completed', 5);
}
