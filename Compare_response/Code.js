function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📝 Text Response")
    .addItem("Generate Token","NLP_token_generate")
    .addItem("Raw data to Normalized","createQuestionsSheetFromRawData")
    .addItem("Get Response from API","processQuestionsFromAPI")
    .addItem("Update column Final Respnse","updateFinalResponse")
    .addItem("Compare Original vs Final Response", "compareTextWithStatusAndReason")    
    .addToUi();
}

