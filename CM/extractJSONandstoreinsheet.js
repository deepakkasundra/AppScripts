// Below function convert department,Ticket schema, Email rules JSON to CSV below is common Function

function extractJSONAndAppendHeaders(data, sheet) {
  try{
  // Dynamically extract all headers from all objects.
  var headersSet = new Set();
  data.forEach(dep => {
    Object.keys(dep).forEach(key => headersSet.add(key));
  });
  var headers = Array.from(headersSet);

  // Append headers to the sheet.
  sheet.appendRow(headers);

  // Map the TicketSchema data into rows and append them to the sheet.
  var rows = data.map(function(dep) {
    return headers.map(function(h) {
      var val = dep[h];
      if (val === null || val === undefined) return "";
      if (typeof val === "object") return JSON.stringify(val); // stringify arrays/objects
      return val;
    });
  });

// Efficiently write all rows at once.
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);

  }
  catch (error) {
    Logger.log("Converting JSON to CSV: " + error.toString());
    SpreadsheetApp.getUi().alert("Converting JSON to CSV: " + error.toString());
    handleError(error);

  }
}

