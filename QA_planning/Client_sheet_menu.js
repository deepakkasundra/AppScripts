function on_open_client() {  
try {

    var LibraryName = "Final_GPT_QA_Report.";  
    var Client_templateId = "18NiMwzrqhZ-oW4ka7v-VDqhaQw21sSm5Zo3hzf9GcME"; 
    
    
    var ui = SpreadsheetApp.getUi();

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var fileId = spreadsheet.getId(); // Get the spreadsheet ID
    var templateId = Client_templateId; // Template file ID
    
    Logger.log("Opened file ID: " + fileId);

    if (fileId === templateId) {
      Logger.log("This is the template. Prompting user to make a copy.");

      ui.alert("⚠️ Please make a copy of this sheet before editing!",
        "📂 Go to File > Make a Copy to create your own version.\n 🚫 GPT QA Menu is not accessible in the template sheet.\n 📩 For assistance, please contact qa_managers@leena.ai or reach out to your manager.",
        ui.ButtonSet.OK
      );

      // Attempt to open the "Make a Copy" page in a new tab
      try {
        var copyUrl = "https://docs.google.com/spreadsheets/d/" + templateId + "/copy";
        Logger.log("Redirecting user to: " + copyUrl);
        
        var html = '<script>window.open("' + copyUrl + '", "_blank"); google.script.host.close();</script>';
        var userInterface = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(100);
        
        ui.showModalDialog(userInterface, "Copy Required");
      }catch (dialogError) {
    Logger.log("Dialog Error: " + dialogError.message);
    // Show a toast message instead of an alert
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ Unable to Load QA Menu in Template sheet.","🛠️ Menu Loader",15);
}


      return; // Stop further execution
    }


    Logger.log("This is a copied version. Proceeding with menu initialization.");






  ui.createMenu('🕵️‍♂️ GPT QA Menu')

.addSubMenu(ui.createMenu('Set Project Type')
.addItem('Set as WorkLM/GPT',LibraryName+'set_workLM')
.addItem('Set as Autonomous',LibraryName+'set_Autonomous'))
    
    .addItem('Validate Policy Name', LibraryName+'checkPolicyNames')
      
    .addSubMenu(ui.createMenu('Create Policy Sheets')
      .addItem('WorkLM/GPT Policy Sheet', LibraryName+'create_sheet_workLM')
            .addItem('Autonomous Policy Sheet',LibraryName+'create_sheet_Autonomous'))

      .addSubMenu(ui.createMenu('Update Formula')
      .addItem('Update WorkLM/GPT Formula', LibraryName+'workLMDynamicFormula')
      .addItem('Update Autonomous Formula', LibraryName+'autonomousDynamicFormula'))
    
    .addItem('Generate Policy HyperLink',LibraryName+'create_sheet_hyper_link')

  //  // .addSeparator()
  //  .addSubMenu(ui.createMenu('Utility')
    //  .addItem('Import Executed Data 1st time only', LibraryName+'mapDataToSheets'))
  //    // .addItem('Delete Unwanted Sheets', 'updateStatusAndDeleteSheets')
     

  .addSubMenu(ui.createMenu('Sync with new Template')
   .addItem('Create WorkLM and Autonomous Template', LibraryName+'CopyWotkLM_Autonomous_Template'))


           .addSubMenu(ui.createMenu('QA Report')
      .addItem('Update QA Cycle Wise Report', LibraryName+'updateQACycleReport')
      .addItem('Consolidated Policy Data for WorkLM/GPT', LibraryName+'Consolidate_sheet_workLM')
      .addItem('Consolidated Policy Data for Autonomous', LibraryName+'Consolidate_sheet_Autonomous')
      
      )
   .addItem('📚 Help', LibraryName+'showSidebar')

    .addToUi();  
  SpreadsheetApp.getActiveSpreadsheet().toast('QA Verification menus Loaded successfully.', '🛠️ Menu Loader', 5);
  }  



// function on_open_client() {
//  try{
//   var LibraryName = "Final_GPT_QA_Report.";
//   SpreadsheetApp.getActiveSpreadsheet().toast('Loading QA Verification Menu. Please wait...', '🛠️ Menu Loader', 5);

// var ui = SpreadsheetApp.getUi();
//   ui.createMenu('🕵️‍♂️ GPT QA Menu')

// //  .addSubMenu(ui.createMenu('Set Work Sheet')
//   //.addItem('Set as WorkLM/GPT',LibraryName+'set_workLM')
//     // .addItem('Set as Autonomous',LibraryName+'set_Autonomous'))
    
//     .addSubMenu(ui.createMenu('Policy Details')
//       .addItem('Validate Policy Name', LibraryName+'checkPolicyNames')
//       .addItem('Create Policy sheets', LibraryName+'create_sheets')
//       .addItem('Generate Policy HyperLink',LibraryName+'create_sheet_hyper_link')
//       .addItem('Update Formula', LibraryName+'setFormulaByDynamicHeader')
//       )
      
//   //  // .addSeparator()
//     .addSubMenu(ui.createMenu('Utility')
//       .addItem('Import Executed Data 1st time only', LibraryName+'mapDataToSheets'))
//   //    // .addItem('Delete Unwanted Sheets', 'updateStatusAndDeleteSheets')
     

//            .addSubMenu(ui.createMenu('QA Report')
//       .addItem('Update QA Cycle Wise Report', LibraryName+'updateQACycleReport')
//       .addItem('Consolidated Policy Data', LibraryName+'consolidateData'))
//    .addItem('📚 Help', LibraryName+'showSidebar')

//     .addToUi();  
//     SpreadsheetApp.getActiveSpreadsheet().toast('QA Verification menus Loaded successfully.', '🛠️ Menu Loader', 5);
//  }

  catch (e) {
handleError(e);
}
}

