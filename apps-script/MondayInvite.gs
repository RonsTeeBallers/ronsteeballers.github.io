function draftMondayGolfInvite() {
  resetMondayResponses();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSheet = ss.getSheetByName("Events");
  var eventData = eventsSheet.getDataRange().getValues();
  var eventRow = null;

  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Monday") {
      eventRow = eventData[i];
      break;
    }
  }

  if (!eventRow) {
    ss.toast("No Monday event found in Events tab.", "Error", 5);
    return;
  }

  var eventDate = eventRow[0];
  var teeTime = eventRow[2];
  var teeTimesReserved = eventRow[3];
  var eventNotes = eventRow[6];

  draftMondayInvite(eventDate, teeTime, teeTimesReserved, eventNotes);
}

function draftMondayInvite(eventDate, teeTime, teeTimesReserved, eventNotes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Format the date
  var formattedDate = (eventDate instanceof Date)
    ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy")
    : eventDate;

  // Format tee time
  var teeTimeStr = (teeTime instanceof Date)
    ? Utilities.formatDate(teeTime, Session.getScriptTimeZone(), "h:mm a")
    : teeTime.toString().substring(0, 5);

  // URLs
  var rsvpLink = "https://docs.google.com/forms/d/e/1FAIpQLSdkxj8VfdqP8Souqn-GtbtPsAHC4TqvECZ3M5ldIzKvMNv27g/viewform?usp=sharing&ouid=115770858933329175920";
  var signupLink = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQgJhTOhxYRXR8CRkd4fJYKX8SenyaRnolamuwgU4oNFwOmacsjAM5-a8QAZ8iT8mIU49DGlCHv0VRM/pubhtml?gid=909010389&single=true";

  // Get Monday player emails
  var mondayGroup = "";

  // Subject
  var subject = "Monday Golf — " + formattedDate + " at Indian Lakes";

  // Body
  var htmlBody =
    "<p>Gentlemen,</p>" +
    "<p>You are invited to join us for golf this Monday. Details below:</p>" +
    "<p>" +
    "<strong>Date:</strong> " + formattedDate + "<br>" +
    "<strong>Course:</strong> Indian Lakes Golf Course<br>" +
    "<strong>Starting Tee Time:</strong> " + teeTimeStr + "<br>" +
    "<strong>Tee Times Reserved:</strong> " + teeTimesReserved +
    "</p>" +
    "<br><p><a href='" + rsvpLink + "' style='background-color:#4CAF50; color:white; padding:10px 20px; text-decoration:none; border-radius:5px; font-weight:bold;'>Click Here to RSVP</a></p><br>" +
    (eventNotes ? "<p><strong>Additional Notes:</strong> " + eventNotes + "</p>" : "") +
    "<p><a href='" + signupLink + "'>See who has signed up as of now</a></p>" +
    "<p>Please RSVP by Saturday.</p>" +
    "<p>See you on the course!<br>Ron</p>";

  GmailApp.createDraft("", subject, htmlBody, {htmlBody: htmlBody});
  ss.toast("Monday draft created! Open Gmail to review and send.", "Golf Invite", 5);
}

function processMondayForm(eventDate, teeTime, teeTimesReserved, eventNotes) {
  resetMondayResponses();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsSheet = ss.getSheetByName("Events");
  eventsSheet.appendRow([
    eventDate,
    'Indian Lakes',
    teeTime,
    teeTimesReserved,
    'Monday',
    'Open',
    eventNotes
  ]);

  draftMondayInvite(eventDate, teeTime, teeTimesReserved, eventNotes);
}

function resetMondayResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Monday Form Responses");
  var lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    // Remove blank rows
    for (var i = lastRow; i >= 2; i--) {
      if (sheet.getRange(i, 1).getValue() === "") {
        sheet.deleteRow(i);
      }
    }
    ss.toast("Monday responses cleared.", "Reset Complete", 5);
  } else {
    ss.toast("No Monday responses to clear.", "Reset", 5);
  }
}
function prepareMondayRSVPSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("Monday Form Responses");
  var rsvpSheet = ss.getSheetByName("Monday RSVPs");

  // Clear the entire RSVPs sheet except headers
  var lastRow = rsvpSheet.getLastRow();
  if (lastRow > 1) {
    rsvpSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  // Get form responses
  var formData = formSheet.getDataRange().getValues();

  // Get unique confirmed Yes players — last response wins
  var responseMap = {};
  for (var i = 1; i < formData.length; i++) {
    var name = formData[i][1];
    var response = formData[i][2];
    var walkRide = formData[i][3];
    var comments = formData[i][4];
    var timestamp = formData[i][0];
    if (name === "" || response !== "Yes") continue;
    responseMap[name] = {
      walkRide: walkRide || "",
      comments: comments || "",
      timestamp: timestamp
    };
  }

  // Write to Monday RSVPs tab
  var outputRow = 2;
  for (var name in responseMap) {
    rsvpSheet.getRange(outputRow, 1).setValue(name);
    rsvpSheet.getRange(outputRow, 2).setValue(responseMap[name].walkRide);
    rsvpSheet.getRange(outputRow, 3).setValue(responseMap[name].comments);
    rsvpSheet.getRange(outputRow, 4).setValue(responseMap[name].timestamp);
    outputRow++;
  }

  // Format Response Time column as date/time
  if (outputRow > 2) {
    rsvpSheet.getRange(2, 4, outputRow - 2, 1).setNumberFormat("M/d/yyyy H:mm");
  }

  ss.toast("Monday RSVP sheet prepared! " + (outputRow - 2) + " players loaded.", "Ready", 5);
}