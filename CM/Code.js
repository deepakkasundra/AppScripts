	// Required API
  // Google Drive API
	// Google Sheets AP
	function onOpen() {
	  var ui = SpreadsheetApp.getUi();
	  
	  // var qamenu = ui.createMenu('⚙️ QA Verification')
		
    var qamenu = ui.createMenu('🕵️‍♂️ QA Verification')
    .addSubMenu(ui.createMenu('Fetch Dashboard Data')
		  .addItem('1. 🚨 Generate Token', 'loginAndStoreToken')
		   
		  )
	// .addSeparator() // Add separator here
		.addSubMenu(ui.createMenu('Fetch Ticketing Form')
		  .addItem('1. Fetch Ticketing Form from UAT (New CM)', 'fetchFromUAT') //single file || Fetch_Ticketing_form_single_code
		  .addItem('2. 🚨 Fetch Ticketing Form from PROD (New CM)','fetchFromPROD')  //single file || Fetch_Ticketing_form_single_code
		  .addItem('3. 🚨 Fetch Ticketing Form from against Form ID PROD (MONO)', 'Fetch_ticketing_Form_Mono_CM') // 'Fetch_ticketing_Form_FromPROD')
		  .addItem('4. 🚨 Fetch Ticketing Form name and URL from PROD (Mono CM)', 'MONO_fetchTicketingFormDataAndUrls_PROD')
		  .addItem('5. Fetch Ticketing Form name and URL from UAT (Mono CM)', 'MONO_fetchTicketingFormDataAndUrls_UAT')
		)

		.addSubMenu(ui.createMenu('Download Ticketing Form')
		  .addItem('1. Download Ticketing form NEW CM (UAT)','downloadLinksAndSaveToDrive_New_CM_UAT')
      .addItem('2. Download Ticketing form NEW CM (PROD)','downloadLinksAndSaveToDrive_New_CM_PROD')        // single file || Download_ticketing_form
		  .addItem('3. Download Ticketing form MONO CM','downloadLinksAndSaveToDrive_Mono_CM') // single file || Download_ticketing_form
			
      )

		  .addSubMenu(ui.createMenu('Validate Ticketing Forms')
		  .addItem('1. Import All Forms and Validate All Forms', 'importXLSXFromLinksAndValidateAndUpdateStatus')
		  .addItem('2. Check Single Ticketing Forms (Manually)', 'verify_ticketing_form_both_CM'))
	
		// .addSeparator()
		.addSubMenu(ui.createMenu('Seprate Exported Forms')
		  .addItem('1. Extract Ticketing Form Data', 'Seprate_category_subcategory_Employeecode_Extrafields_Optimized'))

		.addSubMenu(ui.createMenu('Ticketing form (Seprated Data) vs Assignee Config')
      .addItem('A) Ticketing forms vs. UAT SLA with Param', 'fetchWithEmployeeCode_UAT')
		  .addItem('B) Ticketing forms vs. UAT SLA without Param', 'fetchWithoutEmployeeCode_UAT')		
		  .addItem('C) Ticketing forms vs. Requirement SLA with Param', 'fetchWithEmployeeCode_REQ')
		  .addItem('D) Ticketing forms vs. Requirement SLA without Param', 'fetchWithoutEmployeeCode_REQ')
		  .addItem('E) Ticketing forms vs. PROD SLA with Param', 'fetchWithEmployeeCode_PROD')
		  .addItem('F) Ticketing forms vs. PROD SLA without Param', 'fetchWithoutEmployeeCode_PROD')
		  .addItem('G) Ticketing forms vs. UAT SLA with Param Sub category Optional ', 'fetchWithEmployeeCode_UAT_subcat_optional')
      .addItem('H) Ticketing forms vs. UAT SLA without Param Sub category Optional ', 'fetchWithoutEmployeeCode_UAT_subcat_optional')
      .addItem('I) Ticketing forms vs. PROD SLA with Param Sub category Optional ', 'fetchWithEmployeeCode_PROD_subcat_optional')
		  .addItem('J) Ticketing forms vs. PROD SLA without Param Sub category Optional ', 'fetchWithoutEmployeeCode_PROD_subcat_optional')
		  .addItem('K) Ticketing forms vs. Requirement SLA with Param Sub category Optional ', 'fetchWithEmployeeCode_REQ_subcat_optional')
		  .addItem('L) Ticketing forms vs. Requirement SLA without Param Sub category Optional ', 'fetchWithoutEmployeeCode_REQ_subcat_optional')
      )
		  

		 .addSubMenu(ui.createMenu('Assignee Config Comparison')
			.addItem('A) Compare Requirment vs. UAT SLA with Param','UAT_vs_Requirement_Compare_WithParam')
			  .addItem('B) Compare Requirment vs. UAT SLA without Param','UAT_vs_Requirement_Compare_WithoutParam')
			.addItem('C) Compare PROD vs. UAT SLA without param', 'UAT_vs_PROD_Compare_WithoutParam')
			.addItem('D) Compare PROD vs. UAT SLA with Param', 'UAT_vs_PROD_Compare_WithParam')
      .addItem('E) Compare PROD vs. Req SLA with Param', 'PROD_vs_Requirement_Compare_WithoutParam')
      .addItem('F) Compare PROD vs. Req SLA without Param', 'PROD_vs_Requirement_Compare_WithParam')
      
		 )
	
			.addSubMenu(ui.createMenu('Fetch Webview URL')
			  .addItem('Check UAT Webview URLs', 'fetch_WebviewURL_UAT')
			  .addItem('🚨 Check PROD Webview URLs', 'fetch_WebviewURL_PROD')
      )

			 .addSubMenu(ui.createMenu('Fetch category from Dashboard')
	   .addItem('A. Fetch Category Master from UAT (New CM)', 'fetch_Category_FromUat')
		  .addItem('B. 🚨 Fetch Category Master from PROD (New CM)', 'fetch_Category_FromProd')   )

			 .addSubMenu(ui.createMenu('Category Master verification')  
		  .addItem('A) Seprated data Vs. UAT Category Master (New CM)', 'compareWithUAT')
		  .addItem('B) Seprated data Vs. PROD Category Master (New CM)', 'compareWithPROD')
		  .addItem('C) PROD Category Master Vs. UAT Category Master (New CM)', 'compareWithPRODvsUAT')
		)
    
		   .addSubMenu(ui.createMenu('Seprate Sub Categories by Semi column') // all in single file || Sub_category_sepration_commonFile
			.addItem('PROD Separated by Sub Categories i.e. ; sep.', 'separateRecordsForPROD')
			.addItem('UAT Separated by Sub Categories i.e. ; sep.', 'separateRecordsForUAT')
			.addItem('Requirement sheet Separated by Sub Categories i.e. ; sep.','separateRecordsForREQ'))


			.addSubMenu(ui.createMenu('Email Ticketing')
			  .addItem('Generate NLP Token', 'NLP_token_generate')
			  .addItem('Triaging Email Ticketing (UAT)', 'Triaging_Email_ticketing_UAT')
			  .addItem('Triaging Email Ticketing (PROD)', 'Triaging_Email_ticketing_PROD')
			  .addItem('Send Email', 'sendEmail')
		  )

		.addSubMenu(ui.createMenu('Other Utilities')
				.addItem('Delete Ticketing form sheets','deleteSheetsBasedOnNames')
				.addItem('Retrive file list and URL from Drive folder','runListFilesInFolderByUrl')
				.addItem('Bulk Replacement in Provided folder URL','replaceMultipleValuesInXLSX')
        .addItem('Conditional Replacement', 'Conditional_replaceMultipleValuesInXLSX')
        .addItem('Open JSON to CSV Converter', 'openJsonToCsvDialog')
	   )
	 
		 // .addSubMenu(ui.createMenu('📚Help')
				.addItem('📚Help', 'showHelpDialog')
        // )
	 
	 .addToUi();

  // Dev Tools Menu
  var devMenu = ui.createMenu('🛠️ Dev Tools')
    .addSubMenu(ui.createMenu('Upload Ticketing Form')
      .addItem('1. Create Ticket Form in UAT from URL', 'Create_ticketing_Form_FromUAT')
      .addItem('2. 🚨 Create Ticket Form in PROD from URL', 'Create_ticketing_Form_FromPROD')
      .addItem('3. Upload single Ticketing forms from Local in UAT', 'uploadLocal_UAT')
      .addItem('4. 🚨 Upload single Ticketing forms from Local in PROD', 'uploadLocal_PROD')
      .addItem('5. Upload Ticketing forms from URL in UAT', 'uploadURL_UAT')
      .addItem('6. 🚨 Upload Ticketing forms from URL in PROD', 'uploadURL_PROD')
    )
.addSubMenu(ui.createMenu('Remove Email from Blacklist')
      .addItem('Remove Email From Blacklist - UAT', 'removeFromBlacklistUAT')
      .addItem('🚨Remove Email From Blacklist - PROD', 'removeFromBlacklistPROD')
)
    .addSubMenu(ui.createMenu('Other Utilities')
      .addItem('Delete Ticketing form sheets', 'deleteSheetsBasedOnNames')
      .addItem('Retrive file list and URL from Drive folder', 'runListFilesInFolderByUrl')
      .addItem('Bulk Replacement in Provided folder URL', 'replaceMultipleValuesInXLSX')
      .addItem('Conditional Replacement', 'Conditional_replaceMultipleValuesInXLSX')
      .addItem('Open JSON to CSV Converter', 'openJsonToCsvDialog')
    )
     .addItem('📚Help', 'showHelpDialog')
    .addToUi();
//var adimnMenu = ui.createMenu('Admin')
//.addItem('Export Menu','menu_fileExport')
//.addToUi();
	}

function showHelpDialog() {
  try {
var htmlContent = loadMenuData();
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Help')
        .setWidth(800)
        .setHeight(600)
        .append(`<script>
          document.getElementById('Overview').innerHTML = \`${htmlContent}\`;
        </script>`); // Overview content pass dynamically;    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Help Menu');
  } catch (error) {
    Logger.log('Error showing Help dialog: ' + error.message);
  }
}

function loadMenuData() {
  try {
    Logger.log("Loading menu data...");  // Log when the function is called
    
    var fileId = '@@@@@@@@@@@@@@';  // Menu.js File
    var file = DriveApp.getFileById(fileId);
    
    Logger.log("File ID: " + fileId);  // Log file ID

    var fileBlob = file.getBlob();
    var fileContent = fileBlob.getDataAsString();
    
    Logger.log("File content successfully retrieved.");  // Log file retrieval success

    var jsonData = JSON.parse(fileContent);
    Logger.log("JSON parsed: " + JSON.stringify(jsonData));  // Log parsed JSON data

    var htmlContent = '';
    
    if (jsonData && jsonData.submenu) {
      Logger.log("Building HTML content from JSON...");  // Log building HTML content
      
      htmlContent += `
        <h3 class="highlight">Overview of Script</h3>
        <p>The script adds a custom menu named "<span class="highlight">⚙️ QA Verification</span>" 
        and "<span class="highlight">🛠️ Dev Tools</span>" to the Google Sheets UI with the following 
        sub-menus and options:</p>
        <ol>`;
        
      jsonData.submenu.forEach(function(submenu) {
        htmlContent += `
          <li><span class="highlight">${submenu.title || 'Untitled Submenu'}</span>
            ${submenu.submenuData && submenu.submenuData.length > 0 ? `
              <ul>
                ${submenu.submenuData.map(submenuData => `
                  ${submenuData.items && submenuData.items.length > 0 ? `
                    ${submenuData.items.map(item => `<li>${item}</li>`).join('')}
                  ` : '<li>No items available</li>'}
                `).join('')}
              </ul>
            ` : ''}
          </li>`;
      });

      htmlContent += `</ol>`;
      Logger.log("HTML content generated.");  // Log HTML generation
    } else {
      htmlContent = `<h3 class="highlight">Overview of Script</h3><p class="error-message">No menu data available.</p>`;
      Logger.log("No menu data available in the JSON.");  // Log absence of menu data
    }

    return htmlContent; // Return formatted HTML

  } catch (error) {
    Logger.log('Error loading file data: ' + error.message);  // Log error message
    throw new Error('Error loading file data: ' + error.message);
  }
}

function logError(errorMessage) {
  Logger.log('Error: ' + errorMessage);
}

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

//common message 

function showProgressToast(ss, message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(message, 'Progress', 5); // Display for 5 seconds
    SpreadsheetApp.flush(); // Ensure the UI updates are pushed out immediately
  } catch (error) {
    Logger.log("Error in showProgressToast: " + error.message);
  }
}
