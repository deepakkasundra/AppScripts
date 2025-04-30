

function showSidebar() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Help')
    .setTitle('Help Menu')
     .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput,'Help Menu');
}

