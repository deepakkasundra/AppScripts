// Test comment
// For test git
function processQuestionsFromAPI() {

   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var mainsheet = spreadsheet.getSheetByName('Main'); // Change 'Main' to your sheet name
  var headersRange = mainsheet.getRange(1, 1, 1, mainsheet.getLastColumn());
  var headersValues = headersRange.getValues()[0];
  var rowIndex = 2; // assuming value is available at row 2
  var nlpTokenColumnIndex = mainsheet.getRange(rowIndex, headersValues.indexOf('NLP Token') + 1).getValue();
  var NLP_URL_value = mainsheet.getRange(rowIndex, headersValues.indexOf('NLP Dashboard') + 1).getValue();
  var Project_ID = mainsheet.getRange(rowIndex, headersValues.indexOf('Project ID') + 1).getValue();
  var environment_det = mainsheet.getRange(rowIndex, headersValues.indexOf('Environment') + 1).getValue();
  var batch_zise = mainsheet.getRange(rowIndex, headersValues.indexOf('BatchSize') + 1).getValue();


Logger.log(nlpTokenColumnIndex);

  if (NLP_URL_value === "#N/A" || NLP_URL_value === "" || Project_ID === "" || nlpTokenColumnIndex === "" ) {
    Logger.log("Either Domain name or Token Or Project ID Not available. Further execution stopped.");
    SpreadsheetApp.getActiveSpreadsheet().toast('Either Domain name or NLP User Name OR Password missing', '⚠️ Further execution stopped.', 10);
    return;
  }

  if (nlpTokenColumnIndex === 0) {
    Logger.log("NLP Token column not found.");
    return;
  }

  const sheetName = 'Questions';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const ui = SpreadsheetApp.getUi();

// uat, staging, production,

// No of records can be proced in on batch
  const BATCH_SIZE = batch_zise;
  if (!sheet) return ui.alert(`Sheet "${sheetName}" not found.`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const questionIndex = headers.indexOf('Question');
  if (questionIndex === -1) return ui.alert("Column 'Question' not found.");

  SpreadsheetApp.getActiveSpreadsheet().toast("Initiating API Connection. with Batch Size " + BATCH_SIZE, "Info", 5);

 try {

  const headerMap = {
    'Response': 'responseIndex',
    'Status of API': 'statusIndex',
    'API Response': 'rawApiIndex',
    'Final Response': 'finalresponse'
  };

  const newHeaders = [...headers];
  let columnIndexes = {};

  for (const key in headerMap) {
    let index = newHeaders.indexOf(key);
    if (index === -1) {
      newHeaders.push(key);
      index = newHeaders.length - 1;
    }
    columnIndexes[headerMap[key]] = index;
  }

  if (newHeaders.length !== headers.length) {
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  }

  Object.values(columnIndexes).forEach(idx => clearColumnData(sheet, idx + 1));

  const questions = sheet.getRange(2, questionIndex + 1, sheet.getLastRow() - 1).getValues();



  for (let i = 0; i < questions.length; i += BATCH_SIZE) {
    const batchQuestions = questions.slice(i, i + BATCH_SIZE);
    const batchRequests = [];

    batchQuestions.forEach((q, idx) => {
      const row = i + idx + 2;
      const question = q[0];

      if (!question) {
        sheet.getRange(row, columnIndexes.statusIndex + 1).setValue("Missing question");
        Logger.log(`Row ${row}: Question is empty`);
        return;
      }

      const payload = {
        id: Project_ID,
        query: question,
        env: environment_det,
        check_flow: false,
        project_segment_id: null,
        channelUserId: ""
      };

      const options = {
Domain: `${Domain}/<REDACTED_PATH>/
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        headers: {
          'Authorization': `Bearer ${nlpTokenColumnIndex}`,
          'Accept': '*/*'
        },
        muteHttpExceptions: true
      };
// Logger.log(`URL being called: ${options.url}`);

      batchRequests.push(options);
    });

    const responses = UrlFetchApp.fetchAll(batchRequests);

    responses.forEach((res, idx) => {
      const row = i + idx + 2;
      let reply = "No valid or Longer response in API.";
      const statusCode = res.getResponseCode();
      const body = res.getContentText();

      Logger.log(`Row ${row}: Status Code = ${statusCode}`);
      Logger.log(`Row ${row}: Response Body = ${body}`);

      if (statusCode === 200) {
        try {
          const json = JSON.parse(body);
          const responseText = json?.predictions?.intent_response?.responses?.[0];
          if (responseText) {
            reply = stripHTML(responseText);
          }
          sheet.getRange(row, columnIndexes.responseIndex + 1).setValue(reply);
          sheet.getRange(row, columnIndexes.statusIndex + 1).setValue("Success");
        } catch (parseErr) {
          sheet.getRange(row, columnIndexes.responseIndex + 1).setValue("Parse Error");
          sheet.getRange(row, columnIndexes.statusIndex + 1).setValue("Invalid JSON");
        }
      } else {
        sheet.getRange(row, columnIndexes.responseIndex + 1).setValue("Error");
        sheet.getRange(row, columnIndexes.statusIndex + 1).setValue(`Fail - ${statusCode}`);
      }

//      sheet.getRange(row, columnIndexes.rawApiIndex + 1).setValue(body);
  sheet.getRange(row, columnIndexes.rawApiIndex + 1).setValue(truncateText(body));
  Logger.log(`Full API Response (possibly truncated in sheet): ${body}`);


    });

    SpreadsheetApp.flush(); // Force update after each batch

       // Show progress toast
    const processedCount = Math.min(i + BATCH_SIZE, questions.length);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Processed ${processedCount} of ${questions.length} questions...`, 
      "Progress", 
      5
    );

    Utilities.sleep(300); // Optional: throttle between batches
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Processing completed.", "Done", 5);
  updateFinalResponse();

 }
 catch (error) {
      // Fallback for cases where error parsing fails
handleError(error);
    Logger.log('Error loading file data: ' + error.message);  // Log error message
}
}



function clearColumnData(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, colIndex, lastRow - 1).clearContent();
  }
}

function stripHTML(html) {
  try {
    const output = HtmlService.createHtmlOutput(html).getContent();
    return output
   .replace(/<[^>]*>/g, '')         // Remove HTML tags
      .replace(/&nbsp;/g, ' ')         // Replace non-breaking space
      .replace(/&amp;/g, '&')          // Decode &
      .replace(/&ndash;/g, '-')        // Decode –
      .replace(/&ldquo;/g, '“')        // Decode “
      .replace(/&rdquo;/g, '”')        // Decode ”
      .replace(/&rsquo;/g, '’')        // Decode ’
      .replace(/&lsquo;/g, '‘')        // Decode ‘
      .replace(/&gt;/g, '>');          // Decode >
  } catch (e) {
    Logger.log("Error stripping HTML: " + e.message);
    return html;
  }
}



function truncateText(text, limit = 45000) {
  return (text && text.length > limit) 
    ? text.substring(0, limit) + '... [TRUNCATED]' 
    : text;
}

