/**
 * Compare MONOCM and NewCM sheets based on key, dependsOnKey, dependsOnValue, and options
 */
function MONOCMvsNewCM_form() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monoSheet = ss.getSheetByName("MONOCM"); // MONOCM
  const newcmSheet = ss.getSheetByName("NewCM"); // NewCM

  Logger.log("Starting comparison between 'MONOCM' and 'NewCM' sheets.");

  if (!monoSheet || !newcmSheet) {
    Logger.log("One or both sheets are missing.");
    SpreadsheetApp.getUi().alert("Ensure both 'MONOCM' and 'NewCM' sheets are present.");
    return;
  }

  const monoData = monoSheet.getDataRange().getValues();
  const newcmData = newcmSheet.getDataRange().getValues();

  const monoHeaders = monoData[0];
  const newcmHeaders = newcmData[0];

  const monoKeyIndex = monoHeaders.indexOf("key");
  const monoOptionsIndex = monoHeaders.indexOf("options");
  const monoDependsOnKeyIndex = monoHeaders.indexOf("dependsOnKey");
  const monoDependsOnValueIndex = monoHeaders.indexOf("dependsOnValue");

  const newcmKeyIndex = newcmHeaders.indexOf("key");
  const newcmOptionsIndex = newcmHeaders.indexOf("options");
  const newcmDependsOnKeyIndex = newcmHeaders.indexOf("dependsOnKey");
  const newcmDependsOnValueIndex = newcmHeaders.indexOf("dependsOnValue");

  Logger.log("Header indexes identified for both sheets.");

  if (
    monoKeyIndex === -1 || monoOptionsIndex === -1 || monoDependsOnKeyIndex === -1 || monoDependsOnValueIndex === -1 ||
    newcmKeyIndex === -1 || newcmOptionsIndex === -1 || newcmDependsOnKeyIndex === -1 || newcmDependsOnValueIndex === -1
  ) {
    Logger.log("Missing required columns in one or both sheets.");
    SpreadsheetApp.getUi().alert("Both sheets must have 'key', 'options', 'dependsOnKey', and 'dependsOnValue' columns.");
    return;
  }

  const resultSheetName = "Comparison Results";
  let resultSheet = ss.getSheetByName(resultSheetName);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
    Logger.log("Created new sheet for comparison results.");
  } else {
    resultSheet.clear();
    Logger.log("Cleared existing 'Comparison Results' sheet.");
  }

  resultSheet.appendRow([
    "Key", "Status", "MONOCM Options", "NewCM Options", "MONOCM Original Options", "NewCM Original Options",
    "MONOCM DependsOnValue", "NewCM DependsOnValue", "MONOCM Original DependsOnValue", "NewCM Original DependsOnValue",
    "MONOCM DependsOnKey", "NewCM DependsOnKey", "MONOCM Row", "NewCM Row"
  ]);

  Logger.log("Initialized result sheet with headers.");

  // Create a map of combination keys for NewCM
  const newcmMap = new Map();
  newcmData.slice(1).forEach((row, index) => {
    const newcmKey = row[newcmKeyIndex];
    const newcmDependsOnKey = row[newcmDependsOnKeyIndex] || "";
    const newcmDependsOnValueOriginal = row[newcmDependsOnValueIndex] || "";
    const newcmOptionsOriginal = row[newcmOptionsIndex] || "";

    const newcmDependsOnValue = newcmDependsOnValueOriginal.split("||").map(v => v.trim()).sort().join("||");
    const newcmOptions = newcmOptionsOriginal.split("||").map(v => v.trim()).sort().join("||");

    const combinationKey = `${newcmKey}|${newcmDependsOnKey}|${newcmDependsOnValue}|${newcmOptions}`;
    newcmMap.set(combinationKey, {
      row: index + 2,
      dependsOnKey: newcmDependsOnKey,
      dependsOnValue: newcmDependsOnValue,
      dependsOnValueOriginal: newcmDependsOnValueOriginal,
      options: newcmOptions,
      optionsOriginal: newcmOptionsOriginal
    });

    Logger.log(`NewCM Combination Key=${combinationKey}, Row=${index + 2}`);
  });

  Logger.log("Processed all rows from 'NewCM' sheet.");

  const result = [];

  // Compare MONOCM rows against NewCM map
  monoData.slice(1).forEach((row, index) => {
    const monoKey = row[monoKeyIndex];
    const monoDependsOnKey = row[monoDependsOnKeyIndex] || "";
    const monoDependsOnValueOriginal = row[monoDependsOnValueIndex] || "";
    const monoOptionsOriginal = row[monoOptionsIndex] || "";

    const monoDependsOnValue = monoDependsOnValueOriginal.split("||").map(v => v.trim()).sort().join("||");
    const monoOptions = monoOptionsOriginal.split("||").map(v => v.trim()).sort().join("||");

    const combinationKey = `${monoKey}|${monoDependsOnKey}|${monoDependsOnValue}|${monoOptions}`;
    Logger.log(`MONOCM Combination Key=${combinationKey}, Row=${index + 2}`);

    if (newcmMap.has(combinationKey)) {
      const newcmEntry = newcmMap.get(combinationKey);
      result.push([
        monoKey, "Match",
        monoOptions, newcmEntry.options,
        monoOptionsOriginal, newcmEntry.optionsOriginal,
        monoDependsOnValue, newcmEntry.dependsOnValue,
        monoDependsOnValueOriginal, newcmEntry.dependsOnValueOriginal,
        monoDependsOnKey, newcmEntry.dependsOnKey,
        index + 2, newcmEntry.row
      ]);
      Logger.log(`Key=${monoKey}: Match found in NewCM.`);
    } else {
      result.push([
        monoKey, "No Match",
        monoOptions, "",
        monoOptionsOriginal, "",
        monoDependsOnValue, "",
        monoDependsOnValueOriginal, "",
        monoDependsOnKey, "",
        index + 2, "N/A"
      ]);
      Logger.log(`Key=${monoKey}: No Match found in NewCM.`);
    }
  });

  Logger.log("Processed all rows from 'MONOCM' sheet.");

  // Write results to the sheet
  result.forEach(row => resultSheet.appendRow(row));
  Logger.log("Comparison completed. Results written to the 'Comparison Results' sheet.");
  SpreadsheetApp.getUi().alert("Comparison completed. Check the 'Comparison Results' sheet.");
}

