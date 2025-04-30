// Standard Handle Error message
function handleError(e) {
  // Log the full error details including the stack trace
  Logger.log(`Error in function: ${getFunctionName(e.stack)}, Error: ${e.message}, Stack: ${e.stack}`);
  
  // Create a detailed message to display in the Browser message box
  var detailedMessage = `An error occurred in the function: ${getFunctionName(e.stack)}.\n\nError Details: ${e.message}\n\nPlease contact QA Manager at qa_managers@leena.ai for assistance.\n\nStack Trace:\n${e.stack}`;
  
  // Display the message to the user
  Browser.msgBox(detailedMessage, Browser.Buttons.OK);
}

// Helper function to extract the function name from the stack trace
function getFunctionName(stack) {
  try {
    var functionName = stack.split('at ')[1]; // Extract the function name from the first stack line
    return functionName ? functionName.split(' ')[0] : 'Unknown Function'; // Return the function name or 'Unknown Function'
  } catch (error) {
    return 'Unknown Function'; // If fails to get name, return 'Unknown Function'
  }
}

