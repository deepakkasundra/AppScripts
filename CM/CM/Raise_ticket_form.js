// Function to open Raise Ticket Form for UAT
function openRaiseTicketFormUAT() {
  openRaiseTicketForm('UAT_Category_master');
}

// Function to open Raise Ticket Form for PROD
function openRaiseTicketFormPROD() {
  openRaiseTicketForm('PROD_Category_master');
}

// Main function to open Raise Ticket Form
function openRaiseTicketForm(sheetName) {
  var ui = SpreadsheetApp.getUi();
  
  var departments = fetchDepartments(sheetName); // Fetch unique departments dynamically
  var html = `
    <form>
      <label for="department">Select Department:</label><br>
      <select id="department" name="department" onchange="showLoader(); fetchCategories(this.value)">
        <option value="">Select Department</option>
        ${departments.map(department => `<option value="${department.name}" data-id="${department.id}">${department.name}</option>`).join('')}
      </select><br><br>
      
      <label for="departmentId">Department ID:</label><br>
      <span id="departmentId">N/A</span><br><br>

      <label for="category">Select Category:</label><br>
      <select id="category" name="category" onchange="showLoader(); fetchSubCategories(this.value)">
        <option value="">Select Category</option>
      </select><br><br>
      
      <label for="categoryId">Category ID:</label><br>
      <span id="categoryId">N/A</span><br><br>
      
      <label for="subCategory">Select Sub Category:</label><br>
      <select id="subCategory" name="subCategory" onchange="displaySubCategoryId(this.value)">
        <option value="">Select Sub Category</option>
      </select><br><br>
      
      <label for="subCategoryId">Sub Category ID:</label><br>
      <span id="subCategoryId">N/A</span><br><br>
      
      <div id="loader" style="display: none;">Loading...</div>
      
      <input type="submit" value="Submit">
    </form>
    
    <script>
      var sheetName = "${sheetName}"; // Pass the actual sheet name to JavaScript
      
      function showLoader() {
        document.getElementById('loader').style.display = 'block';
        document.querySelectorAll('select, input').forEach(function(el) {
          el.disabled = true;
        });
      }
      
      function hideLoader() {
        document.getElementById('loader').style.display = 'none';
        document.querySelectorAll('select, input').forEach(function(el) {
          el.disabled = false;
        });
      }
      
      function fetchCategories(departmentName) {
        const departmentId = document.querySelector('option[value="' + departmentName + '"]').dataset.id; // Get selected department ID
        document.getElementById('departmentId').textContent = departmentId; // Display department ID
        google.script.run.withSuccessHandler(function(categories) {
          var categorySelect = document.getElementById('category');
          categorySelect.innerHTML = '<option value="">Select Category</option>';
          document.getElementById('categoryId').textContent = 'N/A'; // Reset category ID
          
          if (categories.length > 0) {
            var uniqueCategories = new Set(); // Ensure unique categories
            categories.forEach(function(category) {
              if (!uniqueCategories.has(category.categoryName)) {
                uniqueCategories.add(category.categoryName);
                var option = document.createElement('option');
                option.value = category.categoryId;  // Use Category ID for form submission
                option.textContent = category.categoryName;  // Display Category Name
                categorySelect.appendChild(option);
              }
            });
          } else {
            categorySelect.innerHTML = '<option value="">No categories available</option>';
          }
          
          // Clear subcategory dropdown and subcategory ID
          document.getElementById('subCategory').innerHTML = '<option value="">Select Sub Category</option>';
          document.getElementById('subCategoryId').textContent = 'N/A';
          
          hideLoader(); // Hide loader after loading categories
        }).fetchCategories(sheetName, departmentName); // Pass the actual sheet name and selected department
      }

      function fetchSubCategories(categoryId) {
        document.getElementById('categoryId').textContent = categoryId; // Display Category ID
        
        // Log the selected category ID for debugging
        console.log('Selected Category ID:', categoryId);
        
        google.script.run.withSuccessHandler(function(subCategories) {
          var subCategorySelect = document.getElementById('subCategory');
          subCategorySelect.innerHTML = '<option value="">Select Sub Category</option>'; // Reset subcategories
          document.getElementById('subCategoryId').textContent = 'N/A'; // Reset subcategory ID

          if (subCategories.length > 0) {
            subCategories.forEach(function(subCategory) {
              var option = document.createElement('option');
              option.value = subCategory.subCategoryId;  // Use Sub Category ID for form submission
              option.textContent = subCategory.subCategoryName;  // Display Sub Category Name
              subCategorySelect.appendChild(option);
            });
          } else {
            subCategorySelect.innerHTML = '<option value="">No subcategories available</option>';
          }
          
          hideLoader(); // Hide loader after loading subcategories
        }).fetchSubCategories(sheetName, categoryId); // Pass the actual sheet name and selected category ID
      }

      function displaySubCategoryId(subCategoryId) {
        document.getElementById('subCategoryId').textContent = subCategoryId; // Display Sub Category ID
      }
    </script>
  `;
  
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(400), 'Raise Ticket');
}


// Function to fetch unique department names dynamically
function fetchDepartments(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  // Get the header row
  var headerRow = data[0]; // First row is the header

  // Dynamically find the indices of the "Department" and "Department ID" columns
  var departmentIndex = headerRow.indexOf("Department");
  var departmentIdIndex = headerRow.indexOf("Department ID");

  // Check if the required columns exist
  if (departmentIndex === -1 || departmentIdIndex === -1) {
    throw new Error("Required columns not found");
  }

  var departments = [];

  // Start from row 1 to skip the header (row index 0)
  for (var i = 1; i < data.length; i++) {
    var departmentName = data[i][departmentIndex]; // Use dynamic index for department name
    var departmentId = data[i][departmentIdIndex]; // Use dynamic index for department ID

    // Add only unique departments
    if (departmentName && !departments.some(dep => dep.name === departmentName)) {
      departments.push({ name: departmentName, id: departmentId }); // Store name and ID
    }
  }

  return departments; // Return array of departments with name and ID
}




// Function to fetch unique categories based on the selected department
function fetchCategories(sheetName, department) {
  Logger.log('Fetching categories for department: ' + department); // Log department for debugging
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = data[0]; // Get the first row as headers

  var departmentIndex = headers.indexOf('Department');
  var categoryIndex = headers.indexOf('Category');
  var categoryIdIndex = headers.indexOf('Category ID');

  if (departmentIndex === -1 || categoryIndex === -1 || categoryIdIndex === -1) {
    Logger.log('Error: Required columns not found');
    return []; // Return empty if columns are not found
  }

  var categoryData = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][departmentIndex] === department) {
      categoryData.push({
        categoryId: data[i][categoryIdIndex],  // Use Category ID as value
        categoryName: data[i][categoryIndex]  // Display Category Name
      });
    }
  }

  Logger.log('Categories found: ' + JSON.stringify(categoryData)); // Log categories found
  return categoryData;
}

// Function to fetch subcategories based on selected category
function fetchSubCategories(sheetName, categoryId) {
  Logger.log('Fetching subcategories for Category ID: ' + categoryId); // Log category ID for debugging
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = data[0]; // Get the first row as headers

  var categoryIdIndex = headers.indexOf('Category ID');
  var subCategoryIndex = headers.indexOf('Sub Categories');
  var subCategoryIdIndex = headers.indexOf('Sub Categories ID');

  if (categoryIdIndex === -1 || subCategoryIndex === -1 || subCategoryIdIndex === -1) {
    Logger.log('Error: Required columns not found');
    return []; // Return empty if columns are not found
  }

  var subCategories = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][categoryIdIndex] === categoryId) {
      subCategories.push({
        subCategoryId: data[i][subCategoryIdIndex],  // Use Sub Category ID as value
        subCategoryName: data[i][subCategoryIndex]   // Display Sub Category Name
      });
    }
  }

  Logger.log('Subcategories found: ' + JSON.stringify(subCategories)); // Log subcategories found
  return subCategories;
}
