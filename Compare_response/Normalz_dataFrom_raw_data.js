function createQuestionsSheetFromRawData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw_data");
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Raw_data' not found.");
    return;
  }

  const data = rawSheet.getDataRange().getValues(); // get all data including headers
  const headers = data[0];
  const identifyIndex = headers.indexOf("Identify");
  const dataIndex = headers.indexOf("Data");

  if (identifyIndex === -1 || dataIndex === -1) {
    SpreadsheetApp.getUi().alert("Columns 'Identify' or 'Data' not found.");
    return;
  }

  const questions = [];

  for (let i = 1; i < data.length - 1; i++) {
    const currentRow = data[i];
    const nextRow = data[i + 1];

    if (currentRow[identifyIndex] === "Q" && nextRow[identifyIndex] === "A") {
      const question = currentRow[dataIndex];
      const answer = nextRow[dataIndex];
      questions.push([question, answer]);
    }
  }

  // Create or clear 'Questions' sheet
  let questionSheet = ss.getSheetByName("Questions");
  if (!questionSheet) {
    questionSheet = ss.insertSheet("Questions");
  } else {
    questionSheet.clearContents();
  }

  // Set headers and data
  questionSheet.getRange(1, 1, 1, 2).setValues([["Question", "Original Text"]]);
  if (questions.length > 0) {
    questionSheet.getRange(2, 1, questions.length, 2).setValues(questions);
  }

  SpreadsheetApp.getUi().alert("Questions sheet created successfully.");
}

