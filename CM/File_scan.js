// Help to get Duplicate and Duplicate row
function findDuplicatesInFolder() {
  

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Folder_Details');
    const headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headersValues = headersRange.getValues()[0];
    const rowIndex = 2; // as value availabe at 2
    const folderUrl = sheet.getRange(rowIndex, headersValues.indexOf('Folder Path (Drive URL)') + 1).getValue();

  Logger.log(folderUrl);
  
  // const folderUrl = 'https://drive.google.com/drive/folders/1ExG9qA8Tbmq3W6LNuMKwHOjRwLPqnUEg'; // Replace with your folder URL
  const folderId = getFolderIdFromUrl(folderUrl);
  
  if (!folderId) {
    Logger.log(`Folder ID could not be extracted from URL: ${folderUrl}`);
    return;
  }
  
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const rptsheet = SpreadsheetApp.getActiveSpreadsheet();
   // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Folder_Details');

  const reportSheetName = 'Duplicate data Report';
  let reportSheet = rptsheet.getSheetByName(reportSheetName);

  // Create report sheet if it doesn't exist
  if (!reportSheet) {
    reportSheet = rptsheet.insertSheet(reportSheetName);
    reportSheet.appendRow(['File Name', 'File URL', 'Cell Location', 'Actual Cell Value', 'Duplicate Value', 'Count', 'Row Duplicate Status', 'Row Position']);
  } else {
    // Clear existing data in the report sheet
    reportSheet.clear();
    reportSheet.appendRow(['File Name', 'File URL', 'Cell Location', 'Actual Cell Value', 'Duplicate Value', 'Count', 'Row Duplicate Status', 'Row Position']);
  }

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const fileName = file.getName();
    const fileUrl = `https://docs.google.com/spreadsheets/d/${fileId}/view`; // Original file URL

    Logger.log(`Processing file: Name = ${fileName}, ID = ${fileId}, URL = ${fileUrl}`);
    
    try {
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        Logger.log(`File is a Google Sheet. Processing directly.`);
        processGoogleSheet(fileId, fileName, fileUrl, reportSheet);
      } else if (file.getMimeType() === MimeType.MICROSOFT_EXCEL) {
        Logger.log(`File is an Excel file. Converting to Google Sheets format.`);
        processExcelFile(fileId, fileName, fileUrl, reportSheet);
      } else {
        Logger.log(`Unsupported file type: ${fileName} with MIME type ${file.getMimeType()}`);
      }
    } catch (e) {
      Logger.log(`Error processing file: ${fileName}, Error: ${e.message}`);
      reportSheet.appendRow([fileName, fileUrl, '', '', '', '', `Error: ${e.message}`, '']);
    }
  }
}

function getFolderIdFromUrl(url) {
  const regex = /[-\w]{25,}/;
  const match = url.match(regex);
  return match ? match[0] : null;
}

function processGoogleSheet(fileId, fileName, fileUrl, reportSheet) {
  try {
    const fileSheet = SpreadsheetApp.openById(fileId);
    const sheets = fileSheet.getSheets();

    sheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      const rowTracker = {};

      data.forEach((row, rowIndex) => {
        const normalizedRow = normalizeRow(row);
        const rowKey = JSON.stringify(normalizedRow);

        if (rowTracker[rowKey]) {
          rowTracker[rowKey].push(rowIndex + 1);
        } else {
          rowTracker[rowKey] = [rowIndex + 1];
        }

        row.forEach((cell, colIndex) => {
          const cellAddress = getColumnLetter(colIndex + 1) + (rowIndex + 1);
          const duplicates = findDuplicatesInCell(cell);

          duplicates.forEach(duplicate => {
            reportSheet.appendRow([fileName, fileUrl, cellAddress, cell, duplicate.value, duplicate.count, '', '']);
          });
        });
      });

      // Report duplicate rows
      for (const [rowKey, indices] of Object.entries(rowTracker)) {
        if (indices.length > 1) {
          const rowPosition = indices.map(index => `Row ${index}`).join(', ');
          reportSheet.appendRow([fileName, fileUrl, '', '', '', 'Duplicate', rowPosition]);
        }
      }

    });
  } catch (e) {
    Logger.log(`Error processing Google Sheet file: ${fileName}, Error: ${e.message}`);
    reportSheet.appendRow([fileName, fileUrl, '', '', '', `Error: ${e.message}`, '']);
  }
}

function processExcelFile(fileId, fileName, fileUrl, reportSheet) {
  try {
    Logger.log(`Attempting to convert Excel file: ${fileName}, ID: ${fileId}`);
    
    // Convert Excel file to Google Sheets format
    const file = DriveApp.getFileById(fileId);
    const convertedFile = Drive.Files.insert(
      {
        mimeType: MimeType.GOOGLE_SHEETS,
        title: file.getName().replace(/\.xlsx$/, '')  // Removing .xlsx extension
      },
      file.getBlob(),
      {
        convert: true  // Ensure the file is converted
      }
    );
    
    const convertedFileId = convertedFile.id;
    const convertedFileUrl = `https://docs.google.com/spreadsheets/d/${convertedFileId}/edit`; // URL for converted file
    
    Logger.log(`Converted file URL: ${convertedFileUrl}`);
    
    // Open the converted file as a Google Sheet
    const spreadsheet = SpreadsheetApp.openById(convertedFileId);
    const sheets = spreadsheet.getSheets();
    
    sheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      const rowTracker = {};

      data.forEach((row, rowIndex) => {
        const normalizedRow = normalizeRow(row);
        const rowKey = JSON.stringify(normalizedRow);

        if (rowTracker[rowKey]) {
          rowTracker[rowKey].push(rowIndex + 1);
        } else {
          rowTracker[rowKey] = [rowIndex + 1];
        }

        row.forEach((cell, colIndex) => {
          const cellAddress = getColumnLetter(colIndex + 1) + (rowIndex + 1);
          const duplicates = findDuplicatesInCell(cell);

          duplicates.forEach(duplicate => {
            reportSheet.appendRow([fileName, fileUrl, cellAddress, cell, duplicate.value, duplicate.count, '', '']);
          });
        });
      });

      // Report duplicate rows
      for (const [rowKey, indices] of Object.entries(rowTracker)) {
        if (indices.length > 1) {
          const rowPosition = indices.map(index => `Row ${index}`).join(', ');
          reportSheet.appendRow([fileName, fileUrl, '', '', '', 'Duplicate', rowPosition]);
        }
      }

    });

    // Delete the converted Google Sheet to avoid clutter
    DriveApp.getFileById(convertedFileId).setTrashed(true);
    Logger.log(`Deleted converted Google Sheet: ${convertedFileId}`);

  } catch (e) {
    Logger.log(`Error processing Excel file: ${fileName}, Error: ${e.message}`);
    reportSheet.appendRow([fileName, fileUrl, '', '', '', `Error: ${e.message}`, '']);
  }
}

function findDuplicatesInCell(cellValue) {
  if (typeof cellValue === 'string') {
    // const values = cellValue.split('||');
        // Normalize the cell value
    const values = cellValue.split('||').map(value => value.trim()); // Trim spaces
    const valueCount = {};
    const duplicates = [];

    values.forEach(value => {
      if (valueCount[value]) {
        valueCount[value]++;
      } else {
        valueCount[value] = 1;
      }
    });

    for (const [value, count] of Object.entries(valueCount)) {
      if (count > 1) {
        duplicates.push({ value, count });
      }
    }

    return duplicates;
  }
  return [];
}

function normalizeRow(row) {
  return row.map(cell => normalizeCell(cell));
}

function normalizeCell(cell) {
  if (typeof cell === 'string') {
    const values = cell.split('||');
    const uniqueValues = [...new Set(values)].sort(); // Remove duplicates and sort
    return uniqueValues.join('||');
  }
  return cell;
}

function getColumnLetter(column) {
  let letter = '';
  while (column > 0) {
    const modulo = (column - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    column = Math.floor((column - modulo) / 26);
  }
  return letter;
}
