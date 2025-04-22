// Script ID - 1ATEILXCu3r9OLpyIaRcIuFJFcqlbaAegGBjxZOPF4pjYrXfIlparG7pv
// Project ID - 94544702572134961226598057
function checkUnusedFunctions() {
  var scriptId = '1ATEILXCu3r9OLpyIaRcIuFJFcqlbaAegGBjxZOPF4pjYrXfIlparG7pv'; // Your script ID
  
  // Fetch the list of files in the script project
  var files = getScriptFiles(scriptId);
  
  // Extract functions and their file names
  var allFunctions = extractFunctions(files);
  
  // Get functions in the onOpen function
  var onOpenFunctions = getFunctionsInOnOpen(files);
  
  // Compare and find unused functions
  var unusedFunctions = allFunctions.map(file => ({
    fileName: file.name,
    unusedFunctions: file.functions.filter(fn => !onOpenFunctions.includes(fn))
  }));
  
  // Log results
  Logger.log('Functions in onOpen:');
  Logger.log(onOpenFunctions.join(', '));
  
  Logger.log('Unused Functions by File:');
  unusedFunctions.forEach(file => {
    if (file.unusedFunctions.length > 0) {
      Logger.log('File Name: ' + file.fileName);
      Logger.log('Unused Functions: ' + file.unusedFunctions.join(', '));
    }
  });
}

function getScriptFiles(scriptId) {
  var files = [];
  
  try {
    var response = UrlFetchApp.fetch(`https://script.googleapis.com/v1/projects/${scriptId}/content`, {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    var data = JSON.parse(response.getContentText());
    
    if (!data.files || data.files.length === 0) {
      Logger.log('No files found in the script project.');
      return files;
    }
    
    data.files.forEach(function(file) {
      if (file.source) {
        files.push({
          name: file.name,
          sourceCode: file.source
        });
        Logger.log('File Name: ' + file.name);
        Logger.log('File Content: ' + file.source);
      }
    });
    
  } catch (error) {
    Logger.log('Error fetching script files: ' + error.message);
  }
  
  return files;
}

function extractFunctions(files) {
  var fileFunctions = [];
  
  files.forEach(function(file) {
    var sourceCode = file.sourceCode;
    var functionNames = [];
    var matches = sourceCode.match(/function\s+([a-zA-Z_]\w*)\s*\(/g);
    
    if (matches) {
      matches.forEach(function(match) {
        var functionName = match.replace(/function\s+|\s*\(/g, '');
        functionNames.push(functionName);
      });
    }
    
    fileFunctions.push({
      name: file.name,
      functions: functionNames
    });
  });
  
  return fileFunctions;
}

function getFunctionsInOnOpen(files) {
  var onOpenCode = getOnOpenCode(files);
  Logger.log('onOpen Code: ' + onOpenCode);
  
  var matches = onOpenCode.match(/addItem\(['"]([^'"]+)['"]\s*,\s*['"]([^'"]+)['"]\)/g);
  
  if (matches) {
    return matches.map(function(match) {
      var functionName = match.match(/addItem\(['"]([^'"]+)['"]\s*,\s*['"]([^'"]+)['"]\)/)[2];
      return functionName;
    });
  }
  
  return [];
}

function getOnOpenCode(files) {
  var onOpenCode = '';
  
  files.forEach(function(file) {
    if (file.name === 'Code') { // Check the correct file name
      onOpenCode = file.sourceCode;
    }
  });
  
  return onOpenCode;
}
