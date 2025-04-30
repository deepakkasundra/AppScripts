function onOpen_1() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🕵️‍♂️ GPT QA Menu')
    .addSubMenu(ui.createMenu('Policy Details')
      .addItem('Validate Policy Name', 'checkPolicyNames')
      .addItem('Create Policy sheets', 'create_sheets')
      .addItem('Generate Policy HyperLink','create_sheet_hyper_link')
      .addItem('Update Formula', 'setFormulaByDynamicHeader')
      )
      
  //  // .addSeparator()
    .addSubMenu(ui.createMenu('Utility')
      .addItem('Import Executed Data 1st time only', 'mapDataToSheets'))
  //    // .addItem('Delete Unwanted Sheets', 'updateStatusAndDeleteSheets')
      .addSubMenu(ui.createMenu('QA Report')
      .addItem('Update QA Cycle Wise Report', 'updateQACycleReport'))
   
   .addItem('📚 Help', 'showSidebar')
   

    .addToUi();

}



// function handleError(e) {
//   // Log the full error details including the stack trace
//   Logger.log(`Error in function: ${getFunctionName(e.stack)}, Error: ${e.message}, Stack: ${e.stack}`);
  
//   // Create a detailed message to display in the Browser message box
//   var detailedMessage = `An error occurred in the function: ${getFunctionName(e.stack)}.\n\nError Details: ${e.message}\n\nPlease contact QA Manager at qa_managers@leena.ai for assistance.\n\nStack Trace:\n${e.stack}`;
  
//   // Display the message to the user
//   Browser.msgBox(detailedMessage, Browser.Buttons.OK);
// }


// Standard Handle Error message
function handleError(e) {
  // Log the full error details including the stack trace
  Logger.log(`Error in function: ${getFunctionName(e.stack)}, Error: ${e.message}, Stack: ${e.stack}`);
  
  // Create a detailed message to display to the user
  var detailedMessage = `An error occurred in the function: ${getFunctionName(e.stack)}.\n\n` +
                        `Error Details: ${e.message}\n\n` +
                        `Please contact QA Manager at qa_managers@leena.ai for assistance.\n\n` +
                        `Stack Trace:\n${e.stack}`;
  
  // Display the message to the user
 SpreadsheetApp.getUi().alert(detailedMessage);
 // Browser.msgBox(detailedMessage, Browser.Buttons.OK);
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


// function to create the index
function createIndex() {
   
  // Get all the different sheet IDs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
   
  var namesArray = sheetNamesIds(sheets);
   
  var indexSheetNames = namesArray[0];
  var indexSheetIds = namesArray[1];
   
  // check if sheet called sheet called already exists
  // if no index sheet exists, create one
  if (ss.getSheetByName('index') == null) 
  {
    var indexSheet = ss.insertSheet('Index',0);    
  }
  // if sheet called index does exist, prompt user for a different name or option to cancel
  else {
     
    var indexNewName = Browser.inputBox('The name Index is already being used, please choose a different name:', 'Please choose another name', Browser.Buttons.OK_CANCEL);
     
    if (indexNewName != 'cancel') {
      var indexSheet = ss.insertSheet(indexNewName,0);
    }
    else {
      Browser.msgBox('No index sheet created');
    }
     
  }
   
  // add sheet title, sheet names and hyperlink formulas
  if (indexSheet) {
     
    printIndex(indexSheet,indexSheetNames,indexSheetIds);
 
  }
     
}
 
// function to update the index, assumes index is the first sheet in the workbook
function updateIndex() 
{   

 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var namesArray = sheetNamesIds(sheets);
  var indexSheetNames = namesArray[0];
  var indexSheetIds = namesArray[1];
   
  // check if sheet called sheet called already exists
  // if no index sheet exists, create one
  if (ss.getSheetByName('index') == null ) 
  {
    var indexNewName = Browser.msgBox('First Create Index through Option available in Menu:',  Browser.Buttons.OK);
   
   }
  // if sheet called index does exist, prompt user for a different name or option to cancel
  
  else {
var idx  =  sheetName(1);

if(idx == "Index")
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var indexSheet = sheets[0];
  var namesArray = sheetNamesIds(sheets);
  var indexSheetNames = namesArray[0];
  var indexSheetIds = namesArray[1];
    
   printIndex(indexSheet,indexSheetNames,indexSheetIds);    

}
else
{
var indexNewName = Browser.msgBox('Please keep "Index" sheet at first position',  Browser.Buttons.OK);

}
 
// Get all the different sheet IDs and print
  
}
}



 
// function to print out the index
function printIndex(sheet,names,formulas) 
{
 
 // sheet.clearContents();

 var ss = SpreadsheetApp.getActiveSpreadsheet();

// clearing existing condition Formating

sheet.getRange("A3:C1000").clearContent();

sheet.getRange("I:I").clearFormat()
sheet.getRange("K:K").clearFormat()




 sheet.getRange(2,2,names.length,1).setValues(names);
 sheet.getRange(2,3,formulas.length,1).setFormulas(formulas);
  
  // Add serial No.
  var LastRow = sheet.getLastRow();
  // Set the first Auto Number
  var AutoNumberStart=0;  
  if (LastRow>1) {
      for(var i=2; i <= LastRow; i++) {
        sheet.getRange(i, 1).setValue(AutoNumberStart);
        AutoNumberStart++;
      }
    }

// addding  header
sheet.getRange(2, 1).setValue("Sr. No.").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(2, 2).setValue("Policy Name").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(2, 3).setValue("Sheet Link").setFontWeight('bold').setBackground("#c9daf8");

sheet.getRange(2, 4).setValue("Status").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(2, 5).setValue("Assigned to").setFontWeight('bold').setBackground("#c9daf8");

sheet.getRange(2, 6).setValue("Test Cases Count").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(2, 7).setValue("Pass").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(3, 7).setValue(`=COUNTIF(INDIRECT("'"&B3&"'!D:D"),$G$2)`);
sheet.getRange(2, 8).setValue("Fail").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(3, 8).setValue(`=COUNTIF(INDIRECT("'"&B3&"'!D:D"),$H$2)`);
sheet.getRange(2, 9).setValue("Not tested").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(3, 9).setValue(`=F3-G3-H3`);
var nottestedRNG = sheet.getRange("I:I");
var rule1 = SpreadsheetApp.newConditionalFormatRule()
.whenNumberGreaterThan(0)
    .setBackground("#f4cccc")
    .setRanges([nottestedRNG])
    .build();

//=COUNTIF(INDIRECT("'"&B3&"'!E:E"),"GPT")
sheet.getRange(2, 10).setValue("Response count from GPT").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(3, 10).setValue(`=COUNTIF(INDIRECT("'"&B3&"'!E:E"),"GPT")`);
sheet.getRange(2, 11).setValue("% response from GPT").setFontWeight('bold').setBackground("#c9daf8");
sheet.getRange(3, 11).setValue(`=if(F3=0,"NA",(J3/F3)*100)`);
var GPTNARNG = sheet.getRange("K:K");
var rule2 = SpreadsheetApp.newConditionalFormatRule()
.whenTextContains("NA")
    .setBackground("#fce5cd")
    .setRanges([GPTNARNG])
    .build();

sheet.getRange(3, 6).setValue(`=COUNTA(INDIRECT("'"&B3&"'!B2:B"))`);

//Header total calculate part
sheet.getRange(1, 6).setValue(`=SUM(F3:F1000)`);
sheet.getRange(1, 7).setValue(`=SUM(G3:G1000)`);
sheet.getRange(1, 8).setValue(`=SUM(H3:H1000)`);
sheet.getRange(1, 9).setValue(`=SUM(I3:I1000)`);
sheet.getRange(1, 10).setValue(`=SUM(J3:J1000)`);
sheet.getRange(1, 11).setValue(`=if(F1=0,"NA",(J1/F1)*100)`);

// var sheet = SpreadsheetApp.getActiveSheet();

//pushing above rule 1 and rule 2
var rules = sheet.getConditionalFormatRules();
rules.push(rule1,rule2);
sheet.setConditionalFormatRules(rules);

}
 
// function to create array of sheet names and sheet ids
function sheetNamesIds(sheets) {
   
  var indexSheetNames = [];
  var indexSheetIds = [];
   
  // create array of sheet names and sheet gids
  sheets.forEach(function(sheet){
    indexSheetNames.push([sheet.getSheetName()]);
    indexSheetIds.push(['=hyperlink("#gid='
                        + sheet.getSheetId() 
                        + '","'
                        + sheet.getSheetName() 
                        + '")']);
  });
   
  return [indexSheetNames, indexSheetIds];
   
}
function countSheets() {
  return SpreadsheetApp.getActive().getSheets().length;
}

function GetSheetName() {
return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function sheetName(idx) {
  if (!idx)
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  else {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var idx = parseInt(idx);
    if (isNaN(idx) || idx < 1 || sheets.length < idx)
      throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
    return sheets[idx-1].getName();
  }
}

// // create sheet 
// function create_sheets() {

// var sheetbyID = sheetName(1);
//  // Get the data from the sheet called CreateSheets
// var ss = SpreadsheetApp.getActiveSpreadsheet();
// var lastRow = ss.getLastRow();
// var sheetNames = SpreadsheetApp.getActive().getSheetByName(sheetbyID).getRange("B2:B"+lastRow).getValues();

//   // For each row in the sheet, insert a new sheet and rename it.
//   sheetNames.forEach(function(row) {
// if (row[0] != "")
// {
// var totalsheet = countSheets();
// var sheetName = row[0];
// if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) != null ) 
//       { 
//               // Browser.msgBox("Duplucate sheet" + sheetName)
//     ss.toast(   "Sheet already exist some of records are skipped " + sheetName, "⚠️ Warning", 5 )
//      }
//                 else
//       {
//           if(sheetName == null || sheetName == '')
//             {
           
//             }
//                 else{
//                 var indexSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName,totalsheet);
//                 // var sheet = SpreadsheetApp.getActive().insertSheet();
//                 // sheet.setName(sheetName);
//   ss.getSheetByName(sheetName).activate;
//   ss.getRange('\'Test Cases Format\'!A1:P100').copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

//       }
//       } 
//   }
//   });

//    ss.toast( ""  , "👍 Process completed ", 5 )


// }


// Update formula 15-05-2022 Finnal changes implemented

// function updateformula(){
// // var ss = SpreadsheetApp.getActiveSheet();

// // var sheet = SpreadsheetApp.getActiveSheet();
// // get active sheet by name
// var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QA_Report");

// sheet.getRange("I2:I").clearFormat()
// sheet.getRange("K2:K").clearFormat()
// // Policy Status
// sheet.getRange(2, 5).setValue(`=IFS(F2=0,"WIP",I2>0,"WIP",K2="NA","WIP",F2=G2+H2,"Done")`);
// // Test Cases Count
// sheet.getRange(2, 6).setValue(`=IF(B2="",0,COUNTA(INDIRECT("'"&B2&"'!B2:B")))`);
// // Pass
// sheet.getRange(2, 7).setValue(`=COUNTIF(INDIRECT("'"&B2&"'!D:D"),$G$1)`);
// // Fail
// sheet.getRange(2, 8).setValue(`=COUNTIF(INDIRECT("'"&B2&"'!D:D"),$H$1)`);
// // Not tested
// sheet.getRange(2, 9).setValue(`=F2-G2-H2`);
// // Response count from GPT
// sheet.getRange(2, 10).setValue(`=COUNTIF(INDIRECT("'"&B2&"'!E:E"),"GPT")`);
// // % response from GPT
// sheet.getRange(2, 11).setValue(`=if(F2=0,"NA",(J2/F2)*100)`);
// // GPT Pass %  = =if(R2=0,"NA",(R2/J2)*100)
// sheet.getRange(2, 12).setValue(`=if(R2=0,"NA",(R2/J2)*100)`);
// // GPT Fail % = =if(S2=0,"0",(S2/J2)*100)
// sheet.getRange(2, 13).setValue(`=if(S2=0,"0",(S2/J2)*100)`);
// // Fallback Count
// sheet.getRange(2, 14).setValue(`=COUNTIF(INDIRECT("'"&B2&"'!E:E"),"Fallback")`);
// // Fallback %
// sheet.getRange(2, 15).setValue(`=if(F2=0,"NA",(N2/F2)*100)`);
// // response from search count
// sheet.getRange(2, 16).setValue(`=COUNTIF(INDIRECT("'"&B2&"'!E:E"),"Search")`);
// // % response from Search
// sheet.getRange(2, 17).setValue(`=if(F2=0,"NA",(P2/F2)*100)`);

// //GPT Pass Count
// // =COUNTIFs(INDIRECT("'"&B2&"'!D:D"),"Pass",INDIRECT("'"&B2&"'!E:E"),"GPT")
// // sheet.getRange(2, 18).setValue(`=COUNTIFs(INDIRECT("'"&B2&"'!D:D"),"Pass",INDIRECT("'"&B2&"'!E:E"),"GPT")`);
// sheet.getRange(2, 18).setValue(`=COUNTIFs(INDIRECT("'"&B2&"'!D:D"),$G$1,INDIRECT("'"&B2&"'!E:E"),"GPT")`);



// //GPT Fail Count
// // =COUNTIFs(INDIRECT("'"&B2&"'!D:D"),"Fail",INDIRECT("'"&B2&"'!E:E"),"GPT")
// //  sheet.getRange(2, 19).setValue(`=COUNTIFs(INDIRECT("'"&B2&"'!D:D"),Fail,INDIRECT("'"&B2&"'!E:E"),"GPT")`);
//  sheet.getRange(2, 19).setValue(`=COUNTIFs(INDIRECT("'"&B2&"'!D:D"),$H$1,INDIRECT("'"&B2&"'!E:E"),"GPT")`);


// // GPT Range condition Format
// var GPTNARNG = sheet.getRange("K:K");
// var rule2 = SpreadsheetApp.newConditionalFormatRule()
// .whenTextContains("NA")
//     .setBackground("#fce5cd")
//     .setRanges([GPTNARNG])
//     .build();

// var nottestedRNG = sheet.getRange("I:I");
// var rule1 = SpreadsheetApp.newConditionalFormatRule()
// .whenNumberGreaterThan(0)
//     .setBackground("#f4cccc")
//     .setRanges([nottestedRNG])
//     .build();
    
// var rules = sheet.getConditionalFormatRules();
// rules.push(rule1,rule2);
// sheet.setConditionalFormatRules(rules);


// SpreadsheetApp.getActiveSpreadsheet().toast("Formula updated for first row please copy it to all records", " 👍Formula Updated" , 5)

// }



// Count total sheet from SS
function countSheets() {
  return SpreadsheetApp.getActive().getSheets().length;
}
// get sheet name from ID
function sheetName(idx) {
  if (!idx)
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  else {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var idx = parseInt(idx);
    if (isNaN(idx) || idx < 1 || sheets.length < idx)
      throw "Invalid parameter (it should be a number from 0 to "+sheets.length+")";
    return sheets[idx-1].getName();
  }

}

// // create hyperlink function
// function create_sheet_hyper_link() {
//   try {  
// var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QA_Report");
//  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
// // Add sheet Link in C 
//   var LastRow = sheet.getLastRow();
//   // Set the first Auto Number
//   var AutoNumberStart=1;  
//   if (LastRow>1) {
//       for(var i=2; i <= LastRow; i++) {
//         // this below line willl add serial numbeer
        
//         sheet.getRange(i, 3).setValue("...Fetching").setFontWeight('bold');
//         sheet.getRange(i, 1).setValue(AutoNumberStart);
// // ss.toast( "Process with Record no " + AutoNumberStart , "⚠️ Warning", 5 )
//        AutoNumberStart++;
//        var rawsheetname = sheet.getRange(i ,2).getValue();
          
//           if(rawsheetname == null || rawsheetname == "")
//             {
              
//              }
//             else
//               {
//         var tmpreadsheetname = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(rawsheetname);
//        if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(rawsheetname) == null)
//         {
//         sheet.getRange(i, 3).setValue("Sheet not found for " + rawsheetname).setFontWeight('Normal');
//         }
//        else
//           {var refSheetId = tmpreadsheetname.getSheetId().toString();
       
//                sheet.getRange(i, 3).setValue(['=hyperlink("#gid='
//                         + refSheetId 
//                         + '","'
//                         + rawsheetname
//                         + '")']).setFontWeight('Normal');              
         
//           }
//               }
//           }
//     }
//     ss.toast( "no. of record processed - " +  sheet.getRange(AutoNumberStart, 1).getValue() , "👍 Process completed ", 5 )
// }
// catch (error) {
//     // Handle errors
//     ss.toast("An error occurred: " + error.message, "Error", 10);
//     Logger.log("Error: " + error.message); // Log the error for debugging purposes
//   }
// }
