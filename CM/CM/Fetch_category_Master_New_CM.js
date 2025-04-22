function fetch_Category_FromProd() {
  fetchData('PROD');
}

function fetch_Category_FromUat() {
  fetchData('UAT');
}

function fetchData(env) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('Main');
    var headersRange = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];
    var rowIndex = 2;

    var botID = mainSheet.getRange(rowIndex, headersValues.indexOf(env + ' BOT ID') + 1).getValue();
    var jwt = mainSheet.getRange(rowIndex, headersValues.indexOf(env + ' JWT') + 1).getValue();
    var dashboardDomain = mainSheet.getRange(rowIndex, headersValues.indexOf('Dashboard Domain Name') + 1).getValue();

    var missingColumns = [];

    if (!botID) {
      missingColumns.push(env + ' BOT ID');
    }

    if (!dashboardDomain) {
      missingColumns.push('Dashboard Domain Name');
    }

    if (missingColumns.length > 0) {
      var errorMessage = "Error: " + missingColumns.join(" and ") + " is missing in " + mainSheet.getName() + " Sheet";
      throw new Error(errorMessage);
    }

    var url = dashboardDomain + '/bots/' + botID + '/@@@@@@@/';

    var headers = {
      'authority': 'staging-case-management-api.leena.ai',
      'accept': 'application/json, text/plain, */*',
      'accept-language': 'en-US,en;q=0.9',
      'authorization': jwt,
      'x-cm-dashboard-user': 'true'
    };

    var options = {
      'headers': headers,
      'muteHttpExceptions': true
    };

    var response;

    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (e) {
      Logger.log('Error fetching data from the API:', e);
      return;
    }

    if (response.getResponseCode() !== 200) {
      Logger.log('Error: Unexpected response code from the API:', response.getResponseCode());
      SpreadsheetApp.getActiveSpreadsheet().toast('Error: Unexpected Error!! response code from the API: ' + response.getResponseCode(), '⚠️ Warning', 10);
      return;
    }

    var responseData = response.getContentText();
    var data;
    try {
      var parsedData = JSON.parse(responseData);
      data = formatData(parsedData.data);
    } catch (e) {
      Logger.log('Error parsing JSON data:', e);
      return;
    }

    if (data.length > 0) {
      var sheet = ss.getSheetByName(env + "_Category_master") || ss.insertSheet(env + "_Category_master");
      sheet.clear();
      ss.setActiveSheet(sheet);
var headerValues = [['Department', 'Department ID', 'Category ID', 'Category', 'categoryCreatedBy', 'Sub Categories ID', 'Sub Categories', 'subCategoryCreatedBy']];

//      var headerValues = [['Department', 'Category ID', 'Category', 'categoryCreatedBy', 'Sub Category ID', 'Sub Categories', 'subCategoryCreatedBy']];
      sheet.getRange(1, 1, 1, headerValues[0].length).setValues(headerValues);

      var values = data.map(function(item) {
        return [
          item.department,
          item.departmentId,
          item.categoryId,
          item.category,
          item.categoryCreatedBy,
          item.subCategoryId,
          item.subCategory,
          item.subCategoryCreatedBy
        ];
      });

      if (values.length > 0) {
        sheet.getRange(2, 1, values.length, values[0].length).setValues(values); // Write data starting from the second row
      }

      Logger.log("Data written to the spreadsheet successfully.");
      showProgressToast(ss, 'Data fetched successfully from ' + env + '!');
    } else {
      Logger.log("No data fetched from the API.");
      showProgressToast(ss, 'No data fetched from the API.');
    }
  } catch (error) {
    Logger.log("Error during Processing: " + error.message);
    ss.toast("Error: " + error.message, 'Error', 10);
  }
}

function formatData(data) {
  const result = [];
  data.forEach((d) => {
    const departments = d.departments.map((dep) => ({
      name: dep.name,
      id: dep._id, // Adding Department ID
    }));
    const subCategories = d.subCategories.map((sub) => {
      return {
        name: sub.name,
        id: sub._id || "",
        createdBy: sub?.createdBy?.displayName || "From Tickets",
      };
    });

    const formattedResult = [];

    departments.forEach((dep) => {
      if (subCategories.length === 0) {
        formattedResult.push({
          department: dep.name,
          departmentId: dep.id, // Adding Department ID
          category: d.name,
          categoryId: d._id, // Adding Category ID
          categoryCreatedBy: d?.createdBy?.displayName || "From Tickets",
          subCategory: "",
          subCategoryId: "", // No Subcategory, leave blank
          subCategoryCreatedBy: "",
        });
      } else {
        subCategories.forEach((sub) => {
          formattedResult.push({
            department: dep.name,
            departmentId: dep.id, // Adding Department ID
            category: d.name,
            categoryId: d._id, // Adding Category ID
            categoryCreatedBy: d?.createdBy?.displayName || "From Tickets",
            subCategory: sub.name,
            subCategoryId: sub.id, // Adding Subcategory ID
            subCategoryCreatedBy: sub.createdBy,
          });
        });
      }
    });
    result.push(...formattedResult);
  });
  return result;
}

// function showProgressToast(ss, message) {
//   ss.toast(message, 'Progress', 5); // Display for 5 seconds
//   SpreadsheetApp.flush(); // Ensure the UI updates are pushed out immediately
// }
