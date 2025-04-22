// check for Menu.js file and restore if deleted.
// check initiated from code.gs file from loadMenuData function.

function getFileOwnerAndLastActivityUser(fileId_menu) {
  try {
    console.log("Script Started");

    const fileId = fileId_menu; 
    if (!fileId) {
      console.error("File ID is missing");
      return;
    }

    console.log("Fetching file from Drive...");
    let file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (e) {
      console.error("Error fetching file: " + e.message);
      return;
    }

    if (!file) {
      console.error("File not found or access denied.");
      return;
    }

    let fileName = file.getName();
    let filePath = getFullFilePath(file);

    console.log("Fetching last activity user...");
    let lastAction = getLastActivityUser(fileId);

    if (lastAction && lastAction.type === "DELETE") {
      console.log(`File was trashed by: ${lastAction.userName} (${lastAction.userEmail})`);
      
      // Restore file
      restoreFile(fileId);

      // Send warning email with details
      if (lastAction.userEmail !== "Email Not Found") {
        sendWarningEmail(lastAction.userEmail, lastAction.userName, fileName, filePath, lastAction.timestamp);
      }
    } else {
      console.log("File was not trashed recently.");
    }

    console.log("Script Completed Successfully");
    return { fileName, filePath, lastUser: lastAction };

  } catch (error) {
    console.error("Unexpected Error: " + error.toString());
  }
}

function getLastActivityUser(fileId) {
  try {
    console.log("Checking Drive Activity API...");

    let request = { "pageSize": 1, "itemName": `items/${fileId}` };
    let response = DriveActivity.Activity.query(request);
    console.log("Drive Activity API response received.");

    if (response.activities && response.activities.length > 0) {
      let activity = response.activities[0];
      
      let isDeleted = activity.primaryActionDetail && activity.primaryActionDetail.delete;
      let timestamp = activity.timestamp || "Timestamp not available";

      let actors = activity.actors;
      for (let actor of actors) {
        if (actor.user && actor.user.knownUser) {
          let userId = actor.user.knownUser.personName.replace("people/", "");
          let userDetails = getUserNameFromId(userId);
          
          return {
            type: isDeleted ? "DELETE" : "OTHER",
            userName: userDetails.name,
            userEmail: userDetails.email,
            timestamp: timestamp
          };
        }
      }
    }
    return null;

  } catch (error) {
    console.error("Error fetching last activity user: " + error.toString());
    return null;
  }
}

function getUserNameFromId(userId) {
  try {
    console.log(`Fetching user details for ID: ${userId}`);

    let person = People.People.get(`people/${userId}`, { personFields: "names,emailAddresses" });

    let fullName = person.names && person.names.length > 0 ? person.names[0].displayName : "User Name Not Found";
    let email = person.emailAddresses && person.emailAddresses.length > 0 ? person.emailAddresses[0].value : "Email Not Found";

    console.log(`User Name: ${fullName}`);
    console.log(`User Email: ${email}`);

    return { name: fullName, email: email };

  } catch (error) {
    console.error(`Error fetching user name: ${error.toString()}`);
    return { name: "Error fetching user details", email: "Email Not Found" };
  }
}

function restoreFile(fileId) {
  try {
    console.log(`Restoring file with ID: ${fileId}`);
    
    let file = DriveApp.getFileById(fileId);
    if (file.isTrashed()) {
      file.setTrashed(false);
      console.log("File successfully restored from trash.");
    } else {
      console.log("File is not in trash, no need to restore.");
    }
  } catch (error) {
    console.error("Error restoring file: " + error.toString());
  }
}

function sendWarningEmail(userEmail, userName, fileName, filePath, trashedDateTime, ccEmails ) {
  try {
    console.log(`Sending warning email to ${userEmail}...`);

    // Convert timestamp to readable format
    let formattedDateTime = formatDateTime(trashedDateTime);

    let subject = `⚠️ WARNING: Unauthorized File Deletion Attempt Detected - ${fileName}`;
    let body = `Dear ${userName},

We detected that the file "${fileName}" was moved to the trash. If this action was unintentional, please be more careful.

**WARNING:** This file has been automatically restored.  
**Unauthorized deletion attempts may be reported.**

File Details:  
File Name: ${fileName}  
File Path: ${filePath}  
Trashed Date & Time: ${formattedDateTime}  

If you believe this was a mistake, please be careful next time.

Best,  
New CM Dynamic Script`;


   let ccUsers = ["qa_managers@leena.ai"];

    // Ensure ccEmails is a valid array
   let emailOptions = { 
      cc: ccUsers.join(",") // Convert array to a comma-separated string
    };
    GmailApp.sendEmail(userEmail, subject, body, emailOptions);
    
    console.log("Warning email sent successfully.");
  } catch (error) {
    console.error("Error sending email: " + error.toString());
  }
}

function formatDateTime(isoDateTime) {
  try {
    let dateObj = new Date(isoDateTime);
    let options = { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', timeZoneName: 'short' };
    return dateObj.toLocaleString("en-US", options);
  } catch (error) {
    console.error("Error formatting date: " + error.toString());
    return isoDateTime; // Fallback
  }
}

function getFullFilePath(file) {
  let path = file.getName();
  let parent = file.getParents().hasNext() ? file.getParents().next() : null;

  while (parent) {
    path = parent.getName() + " / " + path;
    parent = parent.getParents().hasNext() ? parent.getParents().next() : null;
  }
  
  return "My Drive / " + path;
}
