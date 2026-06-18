function draftGolfInvite() {
  resetThursdayResponses();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSheet = ss.getSheetByName("Events");
  var eventData = eventsSheet.getDataRange().getValues();
  var eventRow = null;
  
  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Thursday") {
      eventRow = eventData[i];
      break;
    }
  }
  
  if (!eventRow) {
    ss.toast("No Thursday event found in Events tab.", "Error", 5);
    return;
  }

  var eventDate = eventRow[0];
  var courseName = eventRow[1];
  var teeTime = eventRow[2];
  var teeTimesReserved = eventRow[3];
  var eventNotes = eventRow[6];

  draftThursdayInvite(eventDate, courseName, teeTime, teeTimesReserved, eventNotes);
}

function draftThursdayInvite(eventDate, courseName, teeTime, teeTimesReserved, eventNotes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var coursesSheet = ss.getSheetByName("Courses");

  var formattedDate = (eventDate instanceof Date)
    ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy")
    : eventDate;

  var teeTimeStr = (teeTime instanceof Date)
    ? Utilities.formatDate(teeTime, Session.getScriptTimeZone(), "h:mm a")
    : teeTime.toString().substring(0, 5);

  var courseData = coursesSheet.getDataRange().getValues();
  var courseURL = "";
  var greenFee = "";
  var cartFee = "";
  var totalFee = "";
  var courseNotes = "";

  for (var i = 1; i < courseData.length; i++) {
    if (courseData[i][0] === courseName) {
      courseURL = courseData[i][6];
      greenFee = courseData[i][7];
      cartFee = courseData[i][8];
      totalFee = courseData[i][9];
      courseNotes = courseData[i][10];
      break;
    }
  }

  var greenFeeStr = greenFee ? "$" + parseFloat(greenFee).toFixed(2) : "TBD";
  var cartFeeStr = cartFee ? "$" + parseFloat(cartFee).toFixed(2) : "TBD";
  var totalFeeStr = totalFee ? "$" + parseFloat(totalFee).toFixed(2) : "TBD";
  var rsvpLink = "https://docs.google.com/forms/d/e/1FAIpQLSfet7KLlo8b3LiYiW0hD-9d36pXu0sXKGB7G7bgndkhIBnfDA/viewform?usp=sharing&ouid=115770858933329175920";
  var signupLink = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQgJhTOhxYRXR8CRkd4fJYKX8SenyaRnolamuwgU4oNFwOmacsjAM5-a8QAZ8iT8mIU49DGlCHv0VRM/pubhtml?gid=1459143581&single=true";

  var thursdayGroup = "";

  var subject = "Golf Outing — " + formattedDate + " at " + courseName;

  var htmlBody =
    "<p>Gentlemen,</p>" +
    "<p>You are invited to join us for golf this week. Details below:</p>" +
    "<p>" +
    "<strong>Date:</strong> " + formattedDate + "<br>" +
    "<strong>Course:</strong> " + courseName + "<br>" +
    "<strong>Starting Tee Time:</strong> " + teeTimeStr + "<br>" +
    "<strong>Tee Times Reserved:</strong> " + teeTimesReserved +
    "</p>" +
    "<br><p><a href='" + rsvpLink + "' style='background-color:#4CAF50; color:white; padding:10px 20px; text-decoration:none; border-radius:5px; font-weight:bold;'>Click Here to RSVP</a></p><br>" +
    "<p style='margin:0'><strong>Course Details:</strong> <a href='" + courseURL + "'>" + courseName + " Golf Course</a><br>" +
    "&nbsp;&nbsp;Green Fee: " + greenFeeStr + "<br>" +
    "&nbsp;&nbsp;Cart Fee (per rider): " + cartFeeStr + "<br>" +
    "&nbsp;&nbsp;Total (Green + Cart): " + totalFeeStr +
    "</p>" +
    (courseNotes ? "<p><strong>Course Notes:</strong> " + courseNotes + "</p>" : "") +
    (eventNotes ? "<p><strong>Additional Notes:</strong> " + eventNotes + "</p>" : "") +
    "<p><a href='" + signupLink + "'>See who has signed up as of now</a></p>" +
    "<p>Please RSVP by Tuesday at noon.</p>" +
    "<p>See you on the course!<br>Ron</p>";

  GmailApp.createDraft("", subject, htmlBody, {htmlBody: htmlBody});

  ss.toast("Thursday draft created! Open Gmail to review and send.", "Golf Invite", 5);
}

function processThursdayForm(eventDate, courseName, teeTime, teeTimesReserved, eventNotes) {
  resetThursdayResponses();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSheet = ss.getSheetByName("Events");
  eventsSheet.appendRow([
    eventDate,
    courseName,
    teeTime,
    teeTimesReserved,
    'Thursday',
    'Open',
    eventNotes
  ]);

  draftThursdayInvite(eventDate, courseName, teeTime, teeTimesReserved, eventNotes);
}

function resetThursdayResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Thursday Form Responses");
  var lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    // Remove blank rows
    for (var i = lastRow; i >= 2; i--) {
      if (sheet.getRange(i, 1).getValue() === "") {
        sheet.deleteRow(i);
      }
    }
    ss.toast("Thursday responses cleared.", "Reset Complete", 5);
  } else {
    ss.toast("No Thursday responses to clear.", "Reset", 5);
  }
}
function prepareThursdayPairingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName('Form Responses');
  var rsvpSheet = ss.getSheetByName("Thursday RSVPs");
  var playersSheet = ss.getSheetByName("Players");

  // Clear the entire RSVPs sheet except headers
  var lastRow = rsvpSheet.getLastRow();
  if (lastRow > 1) {
    rsvpSheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  // Get form responses
  var formData = formSheet.getDataRange().getValues();

  // Get players data
  var playersData = playersSheet.getDataRange().getValues();

  // Build player details map from Players tab
  var playerMap = {};
for (var i = 1; i < playersData.length; i++) {
    var lastFirst = playersData[i][0];
    if (lastFirst === "" || lastFirst === null || lastFirst === undefined) continue;
    var lastFirstStr = lastFirst.toString().trim();
    if (lastFirstStr === "") continue;
    playerMap[lastFirstStr.toLowerCase()] = {
      scoringPref: playersData[i][9] || "No Preference",
      preferredPartners: playersData[i][7] || ""
    };
  }

  // Get unique confirmed Yes players — last response wins
  var responseMap = {};
  for (var i = 1; i < formData.length; i++) {
    var name = formData[i][1];
    var response = formData[i][2];
    var walkRide = formData[i][3];
    var comments = formData[i][4];
    if (name === "" || response !== "Yes") continue;
    responseMap[name] = {
      walkRide: walkRide || "",
      comments: comments || ""
    };
  }

  // Write to Thursday RSVPs tab
  var outputRow = 2;
  for (var name in responseMap) {
    var playerDetails = playerMap[name.toString().toLowerCase()] || {
      scoringPref: "",
      preferredPartners: ""
    };
    rsvpSheet.getRange(outputRow, 1).setValue("");
    rsvpSheet.getRange(outputRow, 2).setValue(name);
    rsvpSheet.getRange(outputRow, 3).setValue(responseMap[name].walkRide);
    rsvpSheet.getRange(outputRow, 4).setValue(playerDetails.scoringPref);
    rsvpSheet.getRange(outputRow, 5).setValue(responseMap[name].comments);
    rsvpSheet.getRange(outputRow, 6).setValue(playerDetails.preferredPartners);
    outputRow++;
  }

  ss.toast("Thursday pairing sheet prepared! " + (outputRow - 2) + " players loaded.", "Ready", 5);
  
  function getConfirmedPlayers(eventId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName('Thursday Form Responses');
    var playersSheet = ss.getSheetByName('Players');
    var eventsSheet = ss.getSheetByName('Events');

    // Get event details for slot count
    var eventsData = eventsSheet.getDataRange().getValues();
    var slotsReserved = 0;

    for (var i = 1; i < eventsData.length; i++) {
      var eventDate = eventsData[i][0];
      var dateStr = eventDate instanceof Date
        ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : eventDate.toString().trim();
      if (dateStr === eventId) {
        slotsReserved = parseInt(eventsData[i][3]) * 4;
        break;
      }
    }

    // Get players data for scoring preference lookup
    var playersData = playersSheet.getDataRange().getValues();
    var scoringMap = {};
    for (var i = 1; i < playersData.length; i++) {
      var name = playersData[i][0].toString().trim();
      scoringMap[name] = playersData[i][9] ? playersData[i][9].toString().trim() : 'No Preference';
    }

    // Get confirmed players — last response wins
    var formData = formSheet.getDataRange().getValues();
    var responseMap = {};

    for (var i = 1; i < formData.length; i++) {
      var name = formData[i][1].toString().trim();
      var response = formData[i][2].toString().trim();
      var walkRide = formData[i][3].toString().trim();
      if (name === '') continue;
      responseMap[name] = {
        playing: response,
        walkRide: walkRide || 'No preference'
      };
    }

    // Build confirmed list
    var confirmed = [];
    for (var name in responseMap) {
      if (responseMap[name].playing === 'Yes') {
        confirmed.push({
          name: name,
          walkRide: responseMap[name].walkRide,
          scoring: scoringMap[name] || 'No Preference'
        });
      }
    }

    // Sort alphabetically by last name
    confirmed.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });

    // Convert Last, First to First Last for display
    confirmed = confirmed.map(function(p) {
      var parts = p.name.split(',');
      var display = parts.length > 1
        ? parts[1].trim() + ' ' + parts[0].trim()
        : p.name;
      return {
        name: display,
        walkRide: p.walkRide,
        scoring: p.scoring
      };
    });

    return ContentService.createTextOutput(JSON.stringify({
      confirmed: confirmed,
      count: confirmed.length,
      slots: slotsReserved
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
}