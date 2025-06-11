function fetch_Category_FromProd() {
  fetchData('PROD');
}

function fetch_Category_FromUat() {
  fetchData('UAT');
}

function fetchData(env) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get the main sheet data using the getMainSheetData() function
    const { 
      mainSheet,prodBotId, prodJwt, uatBotId, uatJwt,  domainname 
    } = getMainSheetData();

    // Retrieve the necessary details for the specified environment (PROD/UAT)
    var botID = env === 'PROD' ? prodBotId : uatBotId;
    var jwt = env === 'PROD' ? prodJwt : uatJwt;
    var dashboardDomain = domainname; // Using the common dashboard domain


var missingValues = [];

if (!botID) {
  missingValues.push(env + ' BOT ID value');
}

if (!dashboardDomain) {
  missingValues.push('Dashboard Domain Name value');
}

if (missingValues.length > 0) {
  var errorMessage = "Error: " + missingValues.join(" and ") + " is missing or invalid in the 'Main' Sheet";
  Logger.log(errorMessage); // Log the error message for further inspection

  // Show pop-up message in the UI
  const ui = SpreadsheetApp.getUi();
  ui.alert('⚠️ Missing or Invalid Data', errorMessage, ui.ButtonSet.OK);

  // Stop further execution by throwing an error
  return;
}

const endpoints = getApiEndpoints();
Domain Domain = Domain + '/<REDACTED_PATH>/' + Domain  + Domain(Domain, 'Domain Domain');
Domain Domain = Domain + '/<REDACTED_PATH>/' + Domain  + Domain(Domain, 'Domain Domain Domain');
// https://us-east-1-case-management-api.leena.ai/bots/65e86d6697359e1a7d15b81e/cm/department/list?current=1&perPage=1000&select=name
Logger.log(url)

Logger.log("Department URL: " + departmenturl);

let departmentMap = {};
try {
  const departmentResponse = UrlFetchApp.fetch(departmenturl, {
    'headers': {
      'accept': 'application/json',
      'authorization': jwt,
      'x-cm-dashboard-user': 'true'
    },
    'muteHttpExceptions': true
  });

  if (departmentResponse.getResponseCode() === 200) {
    const departmentData = JSON.parse(departmentResponse.getContentText()).data;
    departmentData.forEach(dept => {
      departmentMap[dept._id] = dept.name;
    });
    Logger.log("Department mapping loaded successfully.");
  } else {
    Logger.log("⚠️ Failed to fetch department data. Status: " + departmentResponse.getResponseCode());
  }
} catch (err) {
  Logger.log("⚠️ Error fetching department data: " + err.message);
}

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
    Logger.log("Category API Raw Response: " + responseData);

    var data;
 var parsedData;
try {
  parsedData = JSON.parse(responseData);
  if (!parsedData.data) {
    throw new Error("Missing 'data' field in parsed JSON.");
  }
  data = formatData(parsedData.data, departmentMap);
} catch (e) {
  Logger.log('Error parsing JSON data: ' + e.message);
  Logger.log('Raw JSON: ' + responseData);
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
handleError(error);
  }
}

function formatData(data, departmentMap) {
  try{
  const result = [];
  data.forEach((d) => {
    const departments = d.departments.map((depId) => {
      return {
        name: departmentMap[depId] || depId,
        id: depId
      };
    });

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
          departmentId: dep.id,
          category: d.name,
          categoryId: d._id,
          categoryCreatedBy: d?.createdBy?.displayName || "From Tickets",
          subCategory: "",
          subCategoryId: "",
          subCategoryCreatedBy: "",
        });
      } else {
        subCategories.forEach((sub) => {
          formattedResult.push({
            department: dep.name,
            departmentId: dep.id,
            category: d.name,
            categoryId: d._id,
            categoryCreatedBy: d?.createdBy?.displayName || "From Tickets",
            subCategory: sub.name,
            subCategoryId: sub.id,
            subCategoryCreatedBy: sub.createdBy,
          });
        });
      }
    });

    result.push(...formattedResult);
  });

  return result;
  }
  catch(error)
  {
    handleError(error);
  }
}

// function showProgressToast(ss, message) {
//   ss.toast(message, 'Progress', 5); // Display for 5 seconds
//   SpreadsheetApp.flush(); // Ensure the UI updates are pushed out immediately
// }

