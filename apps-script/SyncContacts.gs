function syncContactGroups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var playersSheet = ss.getSheetByName("Players");
  var playersData = playersSheet.getDataRange().getValues();

  // Build Thursday and Monday email lists from Players tab
  var thursdayEmails = [];
  var mondayEmails = [];

  for (var i = 1; i < playersData.length; i++) {
    var email = playersData[i][3].toString().trim().toLowerCase();
    var active = playersData[i][4].toString().trim();
    var mondayPlayer = playersData[i][5].toString().trim();
    if (email === "" || email.indexOf("@") === -1) continue;
    if (active === "Yes") thursdayEmails.push(email);
    if (mondayPlayer === "Yes") mondayEmails.push(email);
  }

  var token = ScriptApp.getOAuthToken();

  // Get all contact groups
  var groupsUrl = "https://people.googleapis.com/v1/contactGroups?pageSize=200";
  var groupsResponse = UrlFetchApp.fetch(groupsUrl, {
    headers: { Authorization: "Bearer " + token }
  });
  var groupsData = JSON.parse(groupsResponse.getContentText());
  var groups = groupsData.contactGroups || [];

  Logger.log("Found " + groups.length + " contact groups");

  var thursdayGroupResource = null;
  var mondayGroupResource = null;

  groups.forEach(function(group) {
    Logger.log("Group: " + group.formattedName);
    if (group.formattedName === "Thursday Golf Group") thursdayGroupResource = group.resourceName;
    if (group.formattedName === "Monday Golf Group") mondayGroupResource = group.resourceName;
  });

  if (!thursdayGroupResource) {
    ss.toast("Thursday Golf Group not found in Google Contacts.", "Error", 5);
    return;
  }
  if (!mondayGroupResource) {
    ss.toast("Monday Golf Group not found in Google Contacts.", "Error", 5);
    return;
  }

  Logger.log("Thursday group: " + thursdayGroupResource);
  Logger.log("Monday group: " + mondayGroupResource);

  // Get all contacts with email addresses
  var emailToResource = {};
  var connectionsUrl = "https://people.googleapis.com/v1/people/me/connections?pageSize=1000&personFields=emailAddresses";
  
  var hasMore = true;
  var nextPageToken = null;
  
  while (hasMore) {
    var url = connectionsUrl;
    if (nextPageToken) url += "&pageToken=" + nextPageToken;
    
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token }
    });
    var data = JSON.parse(response.getContentText());
    var connections = data.connections || [];
    
    connections.forEach(function(contact) {
      var emails = contact.emailAddresses || [];
      emails.forEach(function(emailObj) {
        emailToResource[emailObj.value.toLowerCase().trim()] = contact.resourceName;
      });
    });
    
    nextPageToken = data.nextPageToken;
    hasMore = !!nextPageToken;
  }

  Logger.log("Total contacts with emails: " + Object.keys(emailToResource).length);

  // Sync Thursday group
  updateContactGroup(thursdayGroupResource, thursdayEmails, emailToResource, token, "Thursday Golf Group");
  
  // Sync Monday group
  updateContactGroup(mondayGroupResource, mondayEmails, emailToResource, token, "Monday Golf Group");

  ss.toast("Contact groups synced successfully!", "Sync Complete", 5);
}

function updateContactGroup(groupResource, emails, emailToResource, token, groupName) {
  // Get current group members
  var groupUrl = "https://people.googleapis.com/v1/" + groupResource + "?maxMembers=500";
  var groupResponse = UrlFetchApp.fetch(groupUrl, {
    headers: { Authorization: "Bearer " + token }
  });
  var groupData = JSON.parse(groupResponse.getContentText());
  var currentMembers = groupData.memberResourceNames || [];

  // Find resource names for our email list
  var newMembers = [];
  var notFound = [];

  emails.forEach(function(email) {
    if (emailToResource[email]) {
      newMembers.push(emailToResource[email]);
    } else {
      notFound.push(email);
    }
  });

  if (notFound.length > 0) {
    Logger.log(groupName + " — not found in contacts: " + notFound.join(", "));
  }

  // Members to remove
  var toRemove = currentMembers.filter(function(m) {
    return newMembers.indexOf(m) === -1;
  });

  // Members to add
  var toAdd = newMembers.filter(function(m) {
    return currentMembers.indexOf(m) === -1;
  });

  var modifyUrl = "https://people.googleapis.com/v1/" + groupResource + "/members:modify";

  if (toRemove.length > 0 || toAdd.length > 0) {
    var payload = {};
    if (toAdd.length > 0) payload.resourceNamesToAdd = toAdd;
    if (toRemove.length > 0) payload.resourceNamesToRemove = toRemove;

    UrlFetchApp.fetch(modifyUrl, {
      method: "post",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload)
    });
  }

  Logger.log(groupName + " — added: " + toAdd.length + " removed: " + toRemove.length + " not found: " + notFound.length);
}