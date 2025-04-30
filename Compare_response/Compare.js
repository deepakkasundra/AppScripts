
function compareParagraphsWithReason(text1, text2) {
  function normalize(text) {
    return (text || "")
      .toLowerCase()
         .replace(/\n/g, " ") // Replace line breaks with space
        .replace(/[\u2013\u2014]/g, "-") // En dash & em dash to hyphen
         .replace(/(\d)\.(?=[A-Z])/g, "$1. ") // Add space after list numbers like 3.Maternity
      .replace(/([.,!?:;"'()\[\]{}-])/g, " $1 ") // Space around punctuation
      .replace(/[-•–]\s+/g, "\n") // normalize bullets
    .replace(/[:\-–—]+/g, '') // remove colons, dashes
      .replace(/\s+/g, " ")
      .trim();

  }

  const norm1 = normalize(text1);
  const norm2 = normalize(text2);

  if (norm1 === "" && norm2 === "") return ["Not Match", "Both Original and Final Response are missing"];
  if (norm1 === "") return ["Not Match", "Original Text is missing"];
  if (norm2 === "") return ["Not Match", "Final Response is missing"];

  if (norm1 === norm2) return ["Match", ""];

  const words1 = norm1.split(" ");
  const words2 = norm2.split(" ");
  const maxLen = Math.max(words1.length, words2.length);

  let mismatches = [];

  for (let i = 0; i < maxLen; i++) {
    if (!words1[i]) {
      mismatches.push(`Extra word in Final: '${words2[i]}'`);
    } else if (!words2[i]) {
      mismatches.push(`Missing word in Final: '${words1[i]}'`);
    } else if (words1[i] !== words2[i]) {
      mismatches.push(`Mismatch at word ${i + 1}: '${words1[i]}' vs '${words2[i]}'`);
    }
  }

  if (mismatches.length > 0) {
    return ["Not Match", mismatches.join(";\n")];
  }

  return ["Match", ""];
}

// // Normalize and compare function
// function compareParagraphsWithReason(text1, text2) {
//   function normalize(text) {
//     return (text || "")
//       .toLowerCase()
//       .replace(/[.,!?:;"'()\[\]{}-]/g, "")
//       .replace(/\s+/g, " ")
//       .trim();
//   }

//   const norm1 = normalize(text1);
//   const norm2 = normalize(text2);

//   if (norm1 === "" && norm2 === "") return ["Not Match", "Both Original and Final Response are missing"];
//   if (norm1 === "") return ["Not Match", "Original Text is missing"];
//   if (norm2 === "") return ["Not Match", "Final Response is missing"];

//   if (norm1 === norm2) return ["Match", ""];

//   const words1 = norm1.split(" ");
//   const words2 = norm2.split(" ");
//   const maxLen = Math.max(words1.length, words2.length);

//   for (let i = 0; i < maxLen; i++) {
//     if (!words1[i]) return ["Not Match", `Extra word in New: '${words2[i]}'`];
//     if (!words2[i]) return ["Not Match", `Missing word in New: '${words1[i]}'`];
//     if (words1[i] !== words2[i]) return ["Not Match", `'${words1[i]}' vs '${words2[i]}'`];
//   }

//   return ["Not Match", "Unknown difference"];
// }

function compareTextWithStatusAndReason() {
   const sheetName = 'Questions';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
 
 // const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colOriginal = headers.indexOf("Original Text");
  const colNew = headers.indexOf("Final Response");

  let missing = [];
  if (colOriginal === -1) missing.push("Original Text");
  if (colNew === -1) missing.push("Final Response");

  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert(
      `❌ Missing column header(s): ${missing.join(", ")}.\nPlease make sure they exist.`
    );
    return;
  }

  // Find or create "Compare Status" and "Reason" columns
  let colStatus = headers.indexOf("Compare Status");
  let colReason = headers.indexOf("Reason");

  const lastCol = headers.length;

  if (colStatus === -1) {
    colStatus = lastCol;
    sheet.getRange(1, colStatus + 1).setValue("Compare Status");
  } else {
    // Clear Status column values
    sheet.getRange(2, colStatus + 1, sheet.getLastRow() - 1).clearContent();
  }

  if (colReason === -1) {
    colReason = colStatus + 1 === lastCol ? colStatus + 1 : lastCol + 1;
    sheet.getRange(1, colReason + 1).setValue("Reason");
  } else {
    // Clear Reason column values
    sheet.getRange(2, colReason + 1, sheet.getLastRow() - 1).clearContent();
  }

  // Process each row
  for (let i = 1; i < data.length; i++) {
    const original = data[i][colOriginal];
    const newText = data[i][colNew];
    const [status, reason] = compareParagraphsWithReason(original, newText);
    sheet.getRange(i + 1, colStatus + 1).setValue(status);
    sheet.getRange(i + 1, colReason + 1).setValue(reason);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Text comparison completed!", "Done", 3);
}


// function onEdit(e) {
//   const sheet = e.source.getActiveSheet();
//   const editedRow = e.range.getRow();
//   const editedCol = e.range.getColumn();

//   const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//   const colOriginal = headers.indexOf("Original Text");
//   const colNew = headers.indexOf("Final Response");
//   let colStatus = headers.indexOf("Compare Status");
//   let colReason = headers.indexOf("Reason");

//   // If required columns are missing, do nothing
//   if (colOriginal === -1 || colNew === -1) return;

//   // Add Status/Reason columns if missing
//   const lastCol = headers.length;
//   if (colStatus === -1) {
//     colStatus = lastCol;
//     sheet.getRange(1, colStatus + 1).setValue("Compare Status");
//   }
//   if (colReason === -1) {
//     colReason = (colStatus === lastCol) ? colStatus + 1 : lastCol + 1;
//     sheet.getRange(1, colReason + 1).setValue("Reason");
//   }

//   // Only trigger when editing Original Text or Final Response
//   if (editedCol - 1 !== colOriginal && editedCol - 1 !== colNew) return;

//   // Get text values for the current row
//   const original = sheet.getRange(editedRow, colOriginal + 1).getValue();
//   const newText = sheet.getRange(editedRow, colNew + 1).getValue();

//   const [status, reason] = compareParagraphsWithReason(original, newText);

//   // Write Status and Reason
//   sheet.getRange(editedRow, colStatus + 1).setValue(status);
//   sheet.getRange(editedRow, colReason + 1).setValue(reason);
// }

