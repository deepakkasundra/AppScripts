	function onOpen_client_sheet() {
try {
	
    var ui = SpreadsheetApp.getUi();

SpreadsheetApp.getActiveSpreadsheet().toast('Loading QA Verification and Dev Tools menus. Please wait...', '🛠️ Menu Loader', 5);

    var LibraryName = "Dynamic_SLA.";
	  	  // var qamenu = ui.createMenu('⚙️ QA Verification')

    var qamenu = ui.createMenu('🕵️‍♂️ QA Verification')
    .addSubMenu(ui.createMenu('Fetch Dashboard Data')
		  .addItem('1. 🚨 Generate Token', LibraryName+'loginAndStoreToken')
		   )
	// .addSeparator() // Add separator here
		.addSubMenu(ui.createMenu('Fetch Ticketing Form')
		  .addItem('1. Fetch Ticketing Form from UAT (New CM)', LibraryName+'fetchFromUAT') //single file || Fetch_Ticketing_form_single_code
		  .addItem('2. 🚨 Fetch Ticketing Form from PROD (New CM)',LibraryName+'fetchFromPROD')  //single file || Fetch_Ticketing_form_single_code
		  .addItem('3. 🚨 Fetch Ticketing Form from against Form ID PROD (MONO)', LibraryName+'Fetch_ticketing_Form_Mono_CM') // 'Fetch_ticketing_Form_FromPROD')
		  .addItem('4. 🚨 Fetch Ticketing Form name and URL from PROD (Mono CM)', LibraryName+'MONO_fetchTicketingFormDataAndUrls_PROD')
		  .addItem('5. Fetch Ticketing Form name and URL from UAT (Mono CM)', LibraryName+'MONO_fetchTicketingFormDataAndUrls_UAT')
		)
		.addSubMenu(ui.createMenu('Download Ticketing Form')
		  .addItem('1. Download Ticketing form NEW CM (UAT)',LibraryName+'downloadLinksAndSaveToDrive_New_CM_UAT')
      .addItem('2. Download Ticketing form NEW CM (PROD)',LibraryName+'downloadLinksAndSaveToDrive_New_CM_PROD')        // single file || Download_ticketing_form
		  .addItem('3. Download Ticketing form MONO CM',LibraryName+'downloadLinksAndSaveToDrive_Mono_CM') // single file || Download_ticketing_form
			 )
    .addSubMenu(ui.createMenu('Validate Ticketing Forms')
		  .addItem('1. Import All Forms and Validate All Forms', LibraryName+'importXLSXFromLinksAndValidateAndUpdateStatus')
		  .addItem('2. Check Single Ticketing Forms (Manually)', LibraryName+'verify_ticketing_form_both_CM'))

		// .addSeparator()
		.addSubMenu(ui.createMenu('Seprate Exported Forms')
		  .addItem('1. Extract Ticketing Form Data', LibraryName+'Seprate_category_subcategory_Employeecode_Extrafields_Optimized'))

		.addSubMenu(ui.createMenu('Ticketing form (Seprated Data) vs Assignee Config')
    .addItem('A) Ticketing forms vs. UAT SLA with Param', LibraryName+'fetchWithEmployeeCode_UAT')
		  .addItem('B) Ticketing forms vs. UAT SLA without Param', LibraryName+'fetchWithoutEmployeeCode_UAT')		
		  .addItem('C) Ticketing forms vs. Requirement SLA with Param', LibraryName+'fetchWithEmployeeCode_REQ')
		  .addItem('D) Ticketing forms vs. Requirement SLA without Param', LibraryName+'fetchWithoutEmployeeCode_REQ')
		  .addItem('E) Ticketing forms vs. PROD SLA with Param', LibraryName+'fetchWithEmployeeCode_PROD')
		  .addItem('F) Ticketing forms vs. PROD SLA without Param', LibraryName+'fetchWithoutEmployeeCode_PROD')
      .addItem('G) Ticketing forms vs. UAT SLA with Param Sub category Optional ', LibraryName+'fetchWithEmployeeCode_UAT_subcat_optional')
      .addItem('H) Ticketing forms vs. UAT SLA without Param Sub category Optional ', LibraryName+'fetchWithoutEmployeeCode_UAT_subcat_optional')
		  .addItem('G) Ticketing forms vs. PROD SLA with Param Sub category Optional ', LibraryName+'fetchWithEmployeeCode_PROD_subcat_optional')
		  .addItem('H) Ticketing forms vs. PROD SLA without Param Sub category Optional ', LibraryName+'fetchWithoutEmployeeCode_PROD_subcat_optional')
		  .addItem('I) Ticketing forms vs. Requirement SLA with Param Sub category Optional ', LibraryName+'fetchWithEmployeeCode_REQ_subcat_optional')
		  .addItem('J) Ticketing forms vs. Requirement SLA without Param Sub category Optional ', LibraryName+'fetchWithoutEmployeeCode_REQ_subcat_optional')
		       
      )
		 .addSubMenu(ui.createMenu('Assignee Config Comparison')
			.addItem('A) Compare Requirment vs. UAT SLA with Param',LibraryName+'UAT_vs_Requirement_Compare_WithParam')
			  .addItem('B) Compare Requirment vs. UAT SLA without Param',LibraryName+'UAT_vs_Requirement_Compare_WithoutParam')
			.addItem('C) Compare PROD vs. UAT SLA without param', LibraryName+'UAT_vs_PROD_Compare_WithoutParam')
			.addItem('D) Compare PROD vs. UAT SLA with Param', LibraryName+'UAT_vs_PROD_Compare_WithParam')
		 .addItem('E) Compare PROD vs. Req SLA with Param', LibraryName+'PROD_vs_Requirement_Compare_WithoutParam')
      .addItem('F) Compare PROD vs. Req SLA without Param', LibraryName+'PROD_vs_Requirement_Compare_WithParam')
     )
	
			.addSubMenu(ui.createMenu('Fetch Webview URL')
			  .addItem('Check UAT Webview URLs', LibraryName+'fetch_WebviewURL_UAT')
			  .addItem('🚨 Check PROD Webview URLs', LibraryName+'fetch_WebviewURL_PROD')
      )


			 .addSubMenu(ui.createMenu('Fetch Category from Dashboard')
	   .addItem('A. Fetch Category Master from UAT (New CM)', LibraryName+'fetch_Category_FromUat')
		  .addItem('B. 🚨 Fetch Category Master from PROD (New CM)', LibraryName+'fetch_Category_FromProd')  
	   )

			 .addSubMenu(ui.createMenu('Category Master verification')  
	  	  .addItem('A) Seprated data Vs. UAT Category Master (New CM)', LibraryName+'compareWithUAT')
		  .addItem('B) Seprated data Vs. PROD Category Master (New CM)', LibraryName+'compareWithPROD')
	  .addItem('C) PROD Category Master Vs. UAT Category Master (New CM)', LibraryName+'compareWithPRODvsUAT')
	
		)
		   .addSubMenu(ui.createMenu('Seprate Sub Categories by Semi column') // all in single file || Sub_category_sepration_commonFile
			.addItem('PROD Separated by Sub Categories i.e. ; sep.', LibraryName+'separateRecordsForPROD')
			.addItem('UAT Separated by Sub Categories i.e. ; sep.', LibraryName+'separateRecordsForUAT')
			.addItem('Requirement sheet Separated by Sub Categories i.e. ; sep.',LibraryName+'separateRecordsForREQ'))


			.addSubMenu(ui.createMenu('Email Ticketing')
			  .addItem('Generate NLP Token', LibraryName+'NLP_token_generate')
			  .addItem('Triaging Email Ticketing (UAT)', LibraryName+'Triaging_Email_ticketing_UAT')
			  .addItem('Triaging Email Ticketing (PROD)', LibraryName+'Triaging_Email_ticketing_PROD')
			  .addItem('Send Email', LibraryName+'sendEmail')
		  )

		.addSubMenu(ui.createMenu('Other Utilities')
				.addItem('Delete Ticketing form sheets',LibraryName+'deleteSheetsBasedOnNames')
				.addItem('Retrive file list and URL from Drive folder',LibraryName+'runListFilesInFolderByUrl')
				.addItem('Bulk Replacement in Provided folder URL',LibraryName+'replaceMultipleValuesInXLSX')
        .addItem('Conditional Replacement', LibraryName+'Conditional_replaceMultipleValuesInXLSX')
        .addItem('Open JSON to CSV Converter', LibraryName+'openJsonToCsvDialog')
	   )
		 // .addSubMenu(ui.createMenu('📚Help')
				.addItem('📚Help', LibraryName+'showHelpDialog')
        // )
	 .addToUi();
  // Dev Tools Menu
  var devMenu = ui.createMenu('🛠️ Dev Tools')
    .addSubMenu(ui.createMenu('Upload Ticketing Form')
      .addItem('1. Create Ticket Form in UAT from URL', LibraryName+'Create_ticketing_Form_FromUAT')
      .addItem('2. 🚨 Create Ticket Form in PROD from URL', LibraryName+'Create_ticketing_Form_FromPROD')
      .addItem('3. Upload single Ticketing forms from Local in UAT', LibraryName+'uploadLocal_UAT')
      .addItem('4. 🚨 Upload single Ticketing forms from Local in PROD', LibraryName+'uploadLocal_PROD')
      .addItem('5. Upload Ticketing forms from URL in UAT', LibraryName+'uploadURL_UAT')
      .addItem('6. 🚨 Upload Ticketing forms from URL in PROD', LibraryName+'uploadURL_PROD')
    )
    .addSubMenu(ui.createMenu('Remove Email from Blacklist ')
      .addItem('Remove Email From Blacklist - UAT', LibraryName+'removeFromBlacklistUAT')
      .addItem('🚨Remove Email From Blacklist - PROD', LibraryName+'removeFromBlacklistPROD')
)

    .addSubMenu(ui.createMenu('Other Utilities')
      .addItem('Delete Ticketing form sheets', LibraryName+'deleteSheetsBasedOnNames')
      .addItem('Retrive file list and URL from Drive folder', LibraryName+'runListFilesInFolderByUrl')
      .addItem('Bulk Replacement in Provided folder URL', LibraryName+'replaceMultipleValuesInXLSX')
      .addItem('Conditional Replacement', LibraryName+'Conditional_replaceMultipleValuesInXLSX')
      .addItem('Open JSON to CSV Converter', LibraryName+'openJsonToCsvDialog')
    )
    .addItem('📚Help', LibraryName+'showHelpDialog')
    .addToUi();
	SpreadsheetApp.getActiveSpreadsheet().toast('QA Verification and Dev Tools menus Loaded successfully.', '🛠️ Menu Loader', 5);

  }
   catch (e) {
handleError(e);
}

  }