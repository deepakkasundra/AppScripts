
function getHeaderIndex_ACLcompr(sheet, headerName) {
  const headers = sheet.getDataRange().getValues()[0];
  const index = headers.indexOf(headerName);
  if (index === -1) {
    throw new Error(`Header "${headerName}" not found in sheet: ${sheet.getName()}`);
  }
  return index;
}

function compareGroupVsACLUsersPROD() {
  compareGroupVsACLUsers('PROD');
}

function compareGroupVsACLUsersUAT() {
  compareGroupVsACLUsers('UAT');
}

function compareGroupVsACLUsers(env) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupSheetName = `Group_${env}`;
  const aclSheetName = `ACL_Users_${env}`;
  const resultSheetName = `Group_vs_ACL_${env}`;

  const groupSheet = ss.getSheetByName(groupSheetName);
  const aclSheet = ss.getSheetByName(aclSheetName);

  if (!groupSheet || !aclSheet) {
    SpreadsheetApp.getUi().alert(`Either Group_"${env}" or ACL_Users_"${env}" sheets for environment "${env}" are missing. Please Fetch Group and ACL before comparison`);
    return;
  }

  try {
    const groupData = groupSheet.getDataRange().getValues();
    const aclData = aclSheet.getDataRange().getValues();

  // ✅ New check: no data in either sheet
  if (groupData.length <= 1 || aclData.length <= 1) {
    SpreadsheetApp.getUi().alert('Please first fetch Group and ACL user. No data available in one or both sheets.');
    return;
  }



    // Header indices
    const groupNameIndex = getHeaderIndex_ACLcompr(groupSheet, 'name');
    const groupMembersIndex = getHeaderIndex_ACLcompr(groupSheet, 'members');

    const aclIdIndex = getHeaderIndex_ACLcompr(aclSheet, 'id');
    const aclDisplayNameIndex = getHeaderIndex_ACLcompr(aclSheet, 'displayName');
    const aclFirstNameIndex = getHeaderIndex_ACLcompr(aclSheet, 'firstName');
    const aclLastNameIndex = getHeaderIndex_ACLcompr(aclSheet, 'lastName');
    const aclEmailIndex = getHeaderIndex_ACLcompr(aclSheet, 'email');

    // Map ACL users
    const aclUserMap = {};
    for (let i = 1; i < aclData.length; i++) {
      const row = aclData[i];
      const id = row[aclIdIndex];
      if (id) {
        aclUserMap[id] = {
          displayName: row[aclDisplayNameIndex] || '',
          firstName: row[aclFirstNameIndex] || '',
          lastName: row[aclLastNameIndex] || '',
          email: row[aclEmailIndex] || '',
          rowIndex: i + 1
        };
      }
    }

    // Setup result sheet
    let resultSheet = ss.getSheetByName(resultSheetName);
    if (!resultSheet) {
      resultSheet = ss.insertSheet(resultSheetName);
    } else {
      resultSheet.clear();
    }

    const resultHeaders = ['Group Name', 'Member ID', 'Display Name', 'First Name', 'Last Name', 'Email', 'Status', `Group_${env} Row`, `ACL_${env} Row`];
    resultSheet.appendRow(resultHeaders);

    // Process rows incrementally
    let output = [];
    const chunkSize = 50;
    let resultRow = 2;

    for (let i = 1; i < groupData.length; i++) {
      const row = groupData[i];
      const groupName = row[groupNameIndex];
      const membersRaw = row[groupMembersIndex];

      let members = [];
      try {
        if (typeof membersRaw === 'string' && membersRaw.trim().startsWith('[')) {
          members = JSON.parse(membersRaw);
        }
      } catch (e) {
        Logger.log(`Invalid JSON at Group row ${i + 1}: ${membersRaw}`);
        continue;
      }

      for (const member of members) {
        const memberId = member?.id;
        if (!memberId) continue;

        const aclUser = aclUserMap[memberId];
        const status = aclUser ? `Available in ACL ${env}` : `Not Available in ACL ${env}`;

        output.push([
          groupName,
          memberId,
          aclUser?.displayName || '',
          aclUser?.firstName || '',
          aclUser?.lastName || '',
          aclUser?.email || '',
          status,
          i + 1,
          aclUser?.rowIndex || ''
        ]);

        // Write every N rows
        if (output.length >= chunkSize) {
          resultSheet.getRange(resultRow, 1, output.length, output[0].length).setValues(output);
          resultRow += output.length;
          output = [];
        }
      }
    }

    // Write remaining rows
    if (output.length > 0) {
      resultSheet.getRange(resultRow, 1, output.length, output[0].length).setValues(output);
    }

    ss.toast(`Group vs ACL (${env}) comparison done!`, 'Done', 5);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
    Logger.log(error.stack);
handleError(error);
  }
}

