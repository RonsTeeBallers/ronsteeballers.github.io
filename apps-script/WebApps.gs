function doGet(e) {
  var params = e.parameter;
  var action = params.action || '';
  var callback = params.callback || '';
  var result;

  var GUARDED = { createEvent: true, sendInviteEmails: true, savePairings: true, broadcastEmail: true, getPlayers: true, organizerSubmitRSVP: true, getPlayerDetail: true, savePlayer: true, findTeeTimeAvailability: true };
  if (GUARDED[action] && !passcodeOk_(params)) {
    result = ContentService.createTextOutput(JSON.stringify({error: 'Unauthorized: invalid organizer passcode'}))
      .setMimeType(ContentService.MimeType.JSON);
  } else if (action === 'getEvent') {
    result = getEventData(params.eventId, params.player);
  } else if (action === 'getConfirmed') {
    result = getConfirmedPlayers(params.eventId);
  } else if (action === 'getOpenEvents') {
    result = getOpenEvents();
  } else if (action === 'getPlayers') {
    result = getPlayers();
  } else if (action === 'organizerSubmitRSVP') {
    result = organizerSubmitRSVP(params);
  } else if (action === 'getPlayerDetail') {
    result = getPlayerDetail(params);
  } else if (action === 'savePlayer') {
    result = savePlayer(params);
  } else if (action === 'getVenues') {
    result = getVenues();
  } else if (action === 'createEvent') {
    result = createEvent(params);
  } else if (action === 'savePairings') {
    result = savePairings(params);
  } else if (action === 'previewPairings') {
    result = previewPairings(params);
  } else if (action === 'getEventInfo') {
    result = getEventInfo(params.eventId);
  } else if (action === 'getStats') {
    result = getStats();
  } else if (action === 'sendInviteEmails') {
    result = sendInviteEmails(params);
  } else if (action === 'previewInvite') {
    result = previewInvite(params);
  } else if (action === 'broadcastPreview') {
    result = broadcastPreview(params);
  } else if (action === 'broadcastEmail') {
    result = broadcastEmail(params);
  } else if (action === 'checkPasscode') {
    result = checkPasscode(params);
  } else if (action === 'submitRSVP') {
    result = submitRSVP(params);
  } else if (action === 'findTeeTimeAvailability') {
    result = findTeeTimeAvailability(params);
  } else {
    result = ContentService.createTextOutput(JSON.stringify({error: 'Unknown action'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (callback) {
    var json = result.getContent();
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return result;
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action || '';

  if (action === 'submitRSVP') {
    return submitRSVP(data);
  }
  if (action === 'uploadImage') {
    if (!passcodeOk_(data)) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Unauthorized'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return uploadImage(data);
  }

  return ContentService.createTextOutput(JSON.stringify({error: 'Unknown action'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// Organizer passcode gate. Set Script Property ORGANIZER_PASSCODE to enable it.
// If the property is not set, the app stays open (no lockout).
function passcodeOk_(params) {
  var pc = PropertiesService.getScriptProperties().getProperty('ORGANIZER_PASSCODE');
  if (!pc) return true;
  return (params.passcode || '') === pc;
}

function checkPasscode(params) {
  var pc = PropertiesService.getScriptProperties().getProperty('ORGANIZER_PASSCODE');
  var required = !!pc;
  var valid = !required || ((params.passcode || '') === pc);
  return ContentService.createTextOutput(JSON.stringify({ required: required, valid: valid }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Normalize an Events/Form Responses date cell to its canonical yyyy-MM-dd
// event-ID string (the sheet stores real Dates in some rows, strings in others).
function eventDateStr_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value ? value.toString().trim() : '';
}

// Find the Events row for an event ID (its date). Returns
// { row: <values array>, rowIndex: <1-based sheet row> } or null.
function findEventRow_(eventId) {
  var eventsData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events').getDataRange().getValues();
  for (var i = 1; i < eventsData.length; i++) {
    if (eventDateStr_(eventsData[i][0]) === eventId) {
      return { row: eventsData[i], rowIndex: i + 1 };
    }
  }
  return null;
}

// Latest Form Responses row per player for one event, decided by the Timestamp
// in column A - so re-sorting the Form Responses tab can never change which
// response wins. Ties or missing timestamps fall back to "later row wins",
// matching the old behavior on an append-ordered sheet.
function latestRowsByPlayer_(formData, eventId) {
  var map = {};  // "Last, First" -> { ts: millis, row: [...] }
  for (var i = 1; i < formData.length; i++) {
    var name = formData[i][1] ? formData[i][1].toString().trim() : '';
    if (!name) continue;
    if (eventDateStr_(formData[i][6]).replace('EVT-', '') !== eventId) continue;
    var ts = (formData[i][0] instanceof Date) ? formData[i][0].getTime() : 0;
    if (!map[name] || ts >= map[name].ts) map[name] = { ts: ts, row: formData[i] };
  }
  return map;
}

function getEventData(eventId, playerSlug) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var playersSheet = ss.getSheetByName('Players');
    var venuesSheet = ss.getSheetByName('Courses');

    // Find player by slug (last-first format converted to slug)
    var playersData = playersSheet.getDataRange().getValues();
    var playerData = null;

    for (var i = 1; i < playersData.length; i++) {
      var lastFirst = playersData[i][0].toString().toLowerCase();
      var slug = lastFirst.replace(/,\s*/g, '-').replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
      if (slug === playerSlug.toLowerCase()) {
        playerData = {
          first: playersData[i][1].toString(),
          last: playersData[i][2].toString(),
          full: playersData[i][0].toString()
        };
        break;
      }
    }

    if (!playerData) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Player not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Find event by ID (date string). Distinguish closed from missing so the
    // RSVP page can show a friendly message once pairings have gone out.
    var found = findEventRow_(eventId);
    if (!found) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (found.row[5] !== 'Open') {
      return ContentService.createTextOutput(JSON.stringify({
        error: 'This outing is closed - pairings have already been sent.',
        closed: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
    var event = {
      date: eventId,
      venueName: found.row[1].toString(),
      time: found.row[2] instanceof Date
        ? Utilities.formatDate(found.row[2], Session.getScriptTimeZone(), 'h:mm a')
        : found.row[2].toString(),
      slotsReserved: found.row[3].toString(),
      mailingList: found.row[4].toString(),
      notes: found.row[6] ? found.row[6].toString() : ''
    };

    // Find venue details
    var venuesData = venuesSheet.getDataRange().getValues();
    var venue = { name: event.venueName, url: '', greenFee: '', cartFee: '', total: '' };

    for (var i = 1; i < venuesData.length; i++) {
      if (venuesData[i][0].toString() === event.venueName) {
        var green = parseFloat(venuesData[i][7]) || 0;
        var cart = parseFloat(venuesData[i][8]) || 0;
        venue = {
          name: venuesData[i][0].toString(),
          url: venuesData[i][6] ? venuesData[i][6].toString() : '',
          greenFee: green > 0 ? '$' + green.toFixed(2) : '',
          cartFee: cart > 0 ? '$' + cart.toFixed(2) : '',
          total: (green + cart) > 0 ? '$' + (green + cart).toFixed(2) : ''
        };
        break;
      }
    }

    var result = {
      player: playerData,
      event: {
        title: 'Golf Outing - ' + event.venueName,
        date: event.date,
        venue: venue.name,
        venueUrl: venue.url,
        time: event.time,
        slots: event.slotsReserved + ' tee times reserved',
        greenFee: venue.greenFee,
        cartFee: venue.cartFee,
        total: venue.total,
        notes: event.notes,
        signupListUrl: 'https://bit.ly/thursdaygolf'
      }
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOpenEvents() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var eventsSheet = ss.getSheetByName('Events');
    var formSheet = ss.getSheetByName('Form Responses');
    var eventsData = eventsSheet.getDataRange().getValues();
    var formData = formSheet.getDataRange().getValues();  // read once, not per event
    var events = [];

    for (var i = 1; i < eventsData.length; i++) {
      var status = eventsData[i][5] ? eventsData[i][5].toString().trim() : '';
      if (status !== 'Open') continue;

      var dateStr = eventDateStr_(eventsData[i][0]);

      var slotsReserved = parseInt(eventsData[i][3]) || 0;
      var totalSlots = slotsReserved * 4;

      // Count confirmed players for this specific event, latest response per
      // player by Timestamp (sort-order independent)
      var responseMap = latestRowsByPlayer_(formData, dateStr);
      var confirmed = 0;
      for (var name in responseMap) {
        var rrow = responseMap[name].row;
        if (rrow[2].toString().trim() === 'Yes') {
          confirmed++;
          if (rrow[7] && rrow[7].toString().trim()) confirmed++;
        }
      }

      events.push({
        id: dateStr,
        venue: eventsData[i][1].toString(),
        date: dateStr,
        time: eventsData[i][2] instanceof Date
          ? Utilities.formatDate(eventsData[i][2], Session.getScriptTimeZone(), 'h:mm a')
          : eventsData[i][2].toString(),
        slotsReserved: slotsReserved,
        totalSlots: totalSlots,
        confirmed: confirmed,
        mailingList: eventsData[i][4] ? eventsData[i][4].toString() : '',
        notes: eventsData[i][6] ? eventsData[i][6].toString() : ''
      });
    }

    return ContentService.createTextOutput(JSON.stringify({events: events}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getVenues() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Courses');
    var data = sheet.getDataRange().getValues();
    var venues = [];
    for (var i = 1; i < data.length; i++) {
      var name = data[i][0].toString().trim();
      if (name !== '') venues.push(name);
    }
    return ContentService.createTextOutput(JSON.stringify({venues: venues}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createEvent(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Events');
    var date = new Date(params.date + 'T12:00:00');

    // Block a second outing on a date that already has one. The event date is the
    // unique event ID across V2 (RSVPs, pairings, reminders all key on it), so a
    // duplicate date would silently collide. params.date is yyyy-MM-dd.
    var newDateStr = params.date.toString().trim();
    var existing = sheet.getDataRange().getValues();
    for (var e = 1; e < existing.length; e++) {
      if (eventDateStr_(existing[e][0]) === newDateStr) {
        var pretty = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
        return ContentService.createTextOutput(JSON.stringify({
          error: 'An outing is already scheduled for ' + pretty + '. Only one outing per date is allowed - edit or delete the existing event first.',
          duplicate: true
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    var timeParts = params.time.split(':');
    var teeTime = new Date(1970, 0, 1, parseInt(timeParts[0]), parseInt(timeParts[1]));
    sheet.appendRow([
      date,
      params.venue || '',
      teeTime,
      parseInt(params.slots) || 0,
      params.mailingList || 'Main Group',
      'Open',
      params.notes || ''
    ]);
    return ContentService.createTextOutput(JSON.stringify({success: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function getConfirmedPlayers(eventId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName('Form Responses');

    var found = findEventRow_(eventId);
    var slotsReserved = found ? (parseInt(found.row[3]) || 0) * 4 : 0;

    var formData = formSheet.getDataRange().getValues();

    // Latest response per player by Timestamp (sort-order independent)
    var latest = latestRowsByPlayer_(formData, eventId);
    var responseMap = {};
    for (var name in latest) {
      var lrow = latest[name].row;
      responseMap[name] = {
        playing: lrow[2].toString().trim(),
        walkRide: lrow[3].toString().trim() || 'No preference',
        scoring: lrow[4] ? lrow[4].toString().trim() : 'No Preference',
        guest: lrow[7] ? lrow[7].toString().trim() : ''
      };
    }

    // RSVP comments intentionally stay in the spreadsheet only - the organizer
    // reads them there while pairing; they are transient and per-event.
    var confirmed = [];
    for (var name in responseMap) {
      if (responseMap[name].playing === 'Yes') {
        confirmed.push({
          name: name,
          walkRide: responseMap[name].walkRide,
          scoring: responseMap[name].scoring || 'No Preference'
        });
      }
    }

    confirmed.sort(function(a, b) { return a.name.localeCompare(b.name); });

    confirmed = confirmed.map(function(p) {
      var parts = p.name.split(',');
      var display = parts.length > 1
        ? parts[1].trim() + ' ' + parts[0].trim()
        : p.name;
      return { name: display, walkRide: p.walkRide, scoring: p.scoring };
    });

    // Append each guest as a confirmed entry (counts toward the max), tied to host
    for (var hostName in responseMap) {
      if (responseMap[hostName].playing === 'Yes' && responseMap[hostName].guest) {
        var hp = hostName.split(',');
        var hostDisplay = hp.length > 1 ? hp[1].trim() + ' ' + hp[0].trim() : hostName;
        confirmed.push({
          name: responseMap[hostName].guest,
          walkRide: 'No preference',
          scoring: 'No Preference',
          isGuest: true,
          guestOf: hostDisplay
        });
      }
    }

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
function savePairings(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var draftSheet = ss.getSheetByName('Draft Pairings');
    var eventsSheet = ss.getSheetByName('Events');
    var historySheet = ss.getSheetByName('Pairing History');

    // Parse foursomes from pipe-delimited string
    // Format: "Name1|WR1,Name2|WR2;Name3|WR3,Name4|WR4"
    var foursomes = params.foursomes.split(';').map(function(fsStr) {
      if (!fsStr) return [];
      return fsStr.split(',').map(function(playerStr) {
        var parts = playerStr.split('|');
        return { name: parts[0], walkRide: parts[1] || 'No preference' };
      });
    });
    var eventId = params.eventId;

    // Get event details
    var found = findEventRow_(eventId);
    if (!found) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var eventRow = found.row;

    var venueName = eventRow[1].toString();
    var eventDate = eventRow[0];
    var formattedDate = eventDate instanceof Date
      ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')
      : eventDate.toString();

    // Get tee time interval from Courses tab
    var coursesSheet = ss.getSheetByName('Courses');
    var courseData = coursesSheet.getDataRange().getValues();
    var teeTimeInterval = 9;
    for (var c = 1; c < courseData.length; c++) {
      if (courseData[c][0].toString() === venueName) {
        var interval = parseInt(courseData[c][10]);
        if (!isNaN(interval) && interval > 0) teeTimeInterval = interval;
        break;
      }
    }

    // Get starting tee time
    var startTime = eventRow[2];
    var startStr = startTime instanceof Date
      ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'h:mm a')
      : startTime.toString();

    // Parse start minutes
    var startMinutes = 0;
    var timeParts = startStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (timeParts) {
      var h = parseInt(timeParts[1]);
      var m = parseInt(timeParts[2]);
      var ap = timeParts[3].toUpperCase();
      if (ap === 'PM' && h !== 12) h += 12;
      if (ap === 'AM' && h === 12) h = 0;
      startMinutes = h * 60 + m;
    }

    // Clear and write Draft Pairings tab
    draftSheet.clearContents();
    var row = 1;

    foursomes.forEach(function(foursome, idx) {
      var fsMinutes = startMinutes + idx * teeTimeInterval;
      var fh = Math.floor(fsMinutes / 60);
      var fm = fsMinutes % 60;
      var fap = fh >= 12 ? 'PM' : 'AM';
      fh = fh % 12 || 12;
      var teeTimeStr = fh + ':' + (fm < 10 ? '0' + fm : fm) + ' ' + fap;

      var headerCell = draftSheet.getRange(row, 1);
      headerCell.setValue('Foursome ' + (idx + 1) + ' - ' + teeTimeStr);
      headerCell.setFontWeight('bold');
      row++;

      foursome.forEach(function(player) {
        var playerCell = draftSheet.getRange(row, 1);
        playerCell.setValue('   ' + player.name + ' (' + player.walkRide + ')');
        playerCell.setFontWeight('normal');
        row++;
      });

      draftSheet.getRange(row, 1).setFontWeight('normal');
      row++;
    });

    // Build email body (shared with previewPairings)
    var htmlBody = buildPairingsEmailHtml_({
      formattedDate: formattedDate, venueName: venueName,
      foursomes: foursomes, startMinutes: startMinutes,
      teeTimeInterval: teeTimeInterval, comment: params.comment || ''
    });

    var subject = 'Golf Pairings - ' + formattedDate + ' at ' + venueName;

    // Build a map of "First Last" (lowercased) -> email from the Players tab
    var playersSheet = ss.getSheetByName('Players');
    var playersData = playersSheet.getDataRange().getValues();
    var emailByName = {};
    for (var pi = 1; pi < playersData.length; pi++) {
      var pFirst = playersData[pi][1].toString().trim();
      var pLast = playersData[pi][2].toString().trim();
      var pEmail = playersData[pi][3].toString().trim();
      if (!pEmail || pEmail.indexOf('@') === -1) continue;
      var key = (pFirst + ' ' + pLast).toLowerCase().trim();
      emailByName[key] = { email: pEmail, name: pFirst };
    }

    // Collect unique recipients from the foursomes
    var recipients = [];
    var seenEmails = {};
    var unmatched = [];
    foursomes.forEach(function(foursome) {
      foursome.forEach(function(player) {
        var lookup = player.name.toLowerCase().trim();
        var match = emailByName[lookup];
        if (match) {
          if (!seenEmails[match.email]) {
            seenEmails[match.email] = true;
            recipients.push(match);
          }
        } else if (player.name.trim() !== '') {
          unmatched.push(player.name);
        }
      });
    });

    // Send the pairings directly to each player via Brevo (bypassing Gmail)
    var sent = 0;
    var sendErrors = [];
    recipients.forEach(function(recipient) {
      var sendResult = sendBrevoEmail_(recipient.email, recipient.name, subject, htmlBody);
      if (sendResult.ok) {
        sent++;
      } else {
        sendErrors.push(recipient.email + ': ' + sendResult.error);
      }
    });

    // If nothing went out at all (e.g. Brevo down / bad API key), leave the
    // event Open and skip the history write so the organizer can simply fix
    // the problem and hit Send again without creating duplicate history rows.
    if (sent === 0 && recipients.length > 0) {
      return ContentService.createTextOutput(JSON.stringify({
        error: 'No emails were sent (' + (sendErrors[0] || 'unknown send failure') + '). The event is still open - try again.',
        errors: sendErrors
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Save to Pairing History
    foursomes.forEach(function(foursome) {
      var historyRow = [formattedDate];
      foursome.forEach(function(player) { historyRow.push(player.name); });
      while (historyRow.length < 5) historyRow.push('');
      historyRow.push(venueName);
      historySheet.appendRow(historyRow);
    });

    // Close the event now that the pairings are out.
    eventsSheet.getRange(found.rowIndex, 6).setValue('Closed');

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      sent: sent,
      total: recipients.length,
      errors: sendErrors,
      unmatched: unmatched
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }

}

// Build the pairings email HTML. Shared by savePairings (real send) and
// previewPairings (no-send preview) so they never drift apart.
// o = { formattedDate, venueName, foursomes, startMinutes, teeTimeInterval, comment }
function buildPairingsEmailHtml_(o) {
  // One styled block per foursome, matching the invite's card look.
  var foursomeBlocks = o.foursomes.map(function(foursome, idx) {
    var fsMinutes = o.startMinutes + idx * o.teeTimeInterval;
    var fh = Math.floor(fsMinutes / 60);
    var fm = fsMinutes % 60;
    var fap = fh >= 12 ? 'PM' : 'AM';
    fh = fh % 12 || 12;
    var teeTimeStr = fh + ':' + (fm < 10 ? '0' + fm : fm) + ' ' + fap;

    var playerRows = foursome.map(function(player) {
      // Guest names originate from the public RSVP form - escape them.
      return '<p style="margin:4px 0;font-size:15px;color:#1a2332;">' +
        escapeHtml_(player.name) + ' <span style="color:#5d6d7e;font-size:13px;">(' + escapeHtml_(player.walkRide) + ')</span></p>';
    }).join('');

    return '<div style="background:#f0f4f8;padding:14px 16px;border-radius:8px;margin:0 0 12px;">' +
      '<p style="margin:0 0 6px;font-size:16px;font-weight:700;color:#1a5276;">&#9971; Foursome ' + (idx + 1) +
      ' <span style="color:#5d6d7e;font-weight:400;font-size:14px;">&mdash; ' + teeTimeStr + '</span></p>' +
      playerRows +
      '</div>';
  }).join('');

  var commentSection = o.comment
    ? '<p style="background:#eaf2ff;padding:12px;border-radius:8px;border-left:3px solid #1a5276;color:#1a2332;font-size:15px;line-height:1.5;">' + o.comment + '</p>'
    : '';

  return '<meta charset="utf-8">' +
    '<div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;max-width:600px;margin:0 auto;padding:20px;">' +
    '<div style="background:#1a5276;color:white;padding:20px;border-radius:12px 12px 0 0;text-align:center;">' +
    '<div style="font-size:32px;">&#9971;</div>' +
    '<h1 style="margin:8px 0;font-size:22px;">Golf Pairings</h1>' +
    '</div>' +
    '<div style="background:white;padding:24px;border:1px solid #e0e8f0;border-top:none;">' +
    '<p style="font-size:18px;font-weight:700;color:#1a2332;">Gentlemen,</p>' +
    '<p style="color:#5d6d7e;">Here are the pairings for ' + o.formattedDate + ' at ' + o.venueName + ':</p>' +
    commentSection +
    foursomeBlocks +
    '<p style="color:#5d6d7e;font-size:14px;">Tee times reserved under the name Ron Blanton.</p>' +
    '</div>' +
    '<div style="background:#f0f4f8;padding:12px;border-radius:0 0 12px 12px;text-align:center;">' +
    '<p style="color:#aab7c4;font-size:12px;margin:0;">See you on the course!</p>' +
    '</div>' +
    '</div>';
}

// Build a no-send preview of the pairings email (mirrors previewInvite).
// Does NOT write Draft Pairings, history, send email, or close the event.
function previewPairings(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var eventId = params.eventId;

    var foursomes = (params.foursomes || '').split(';').map(function(fsStr) {
      if (!fsStr) return [];
      return fsStr.split(',').map(function(playerStr) {
        var parts = playerStr.split('|');
        return { name: parts[0], walkRide: parts[1] || 'No preference' };
      });
    });

    // Find the event row
    var found = findEventRow_(eventId);
    if (!found) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var eventRow = found.row;

    var venueName = eventRow[1].toString();
    var formattedDate = eventRow[0] instanceof Date
      ? Utilities.formatDate(eventRow[0], Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')
      : eventRow[0].toString();

    // Tee time interval from Courses tab
    var coursesSheet = ss.getSheetByName('Courses');
    var courseData = coursesSheet.getDataRange().getValues();
    var teeTimeInterval = 9;
    for (var c = 1; c < courseData.length; c++) {
      if (courseData[c][0].toString() === venueName) {
        var interval = parseInt(courseData[c][10]);
        if (!isNaN(interval) && interval > 0) teeTimeInterval = interval;
        break;
      }
    }

    // Starting tee time -> minutes
    var startTime = eventRow[2];
    var startStr = startTime instanceof Date
      ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'h:mm a')
      : startTime.toString();
    var startMinutes = 0;
    var timeParts = startStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (timeParts) {
      var h = parseInt(timeParts[1]);
      var m = parseInt(timeParts[2]);
      var ap = timeParts[3].toUpperCase();
      if (ap === 'PM' && h !== 12) h += 12;
      if (ap === 'AM' && h === 12) h = 0;
      startMinutes = h * 60 + m;
    }

    var htmlBody = buildPairingsEmailHtml_({
      formattedDate: formattedDate, venueName: venueName,
      foursomes: foursomes, startMinutes: startMinutes,
      teeTimeInterval: teeTimeInterval, comment: params.comment || ''
    });

    // Count unique recipients that map to a Players-tab email (mirrors savePairings)
    var playersSheet = ss.getSheetByName('Players');
    var playersData = playersSheet.getDataRange().getValues();
    var emailByName = {};
    for (var pi = 1; pi < playersData.length; pi++) {
      var pEmail = playersData[pi][3].toString().trim();
      if (!pEmail || pEmail.indexOf('@') === -1) continue;
      var key = (playersData[pi][1].toString().trim() + ' ' + playersData[pi][2].toString().trim()).toLowerCase().trim();
      emailByName[key] = true;
    }
    var seen = {};
    var count = 0;
    var unmatched = [];
    foursomes.forEach(function(foursome) {
      foursome.forEach(function(player) {
        var lookup = player.name.toLowerCase().trim();
        if (emailByName[lookup]) {
          if (!seen[lookup]) { seen[lookup] = true; count++; }
        } else if (player.name.trim() !== '') {
          unmatched.push(player.name);
        }
      });
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      html: htmlBody,
      subject: 'Golf Pairings - ' + formattedDate + ' at ' + venueName,
      count: count,
      unmatched: unmatched
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Shared helper: send one transactional email through Brevo.
// Returns { ok: true } on success, or { ok: false, error: '...' } on failure.
function sendBrevoEmail_(toEmail, toName, subject, htmlBody) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'Brevo API key not configured (Script Property BREVO_API_KEY)' };
  }
  try {
    var response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
      method: 'post',
      headers: {
        'api-key': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        sender: { name: 'RonsTeeBallers', email: 'ronsteeballers@gmail.com' },
        replyTo: { name: 'RonsTeeBallers', email: 'ronsteeballers@gmail.com' },
        to: [{ email: toEmail, name: toName || toEmail }],
        subject: subject,
        htmlContent: htmlBody
      }),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code === 200 || code === 201) {
      return { ok: true };
    }
    return { ok: false, error: 'HTTP ' + code + ' ' + response.getContentText() };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// Run this manually from the editor to confirm Brevo sends a pairings-style
// email to yourself. Check the Execution log for the result.
function testBrevoSend() {
  var result = sendBrevoEmail_(
    'ronsteeballers@gmail.com',
    'Ron',
    'TEST - Golf Pairings email',
    '<p>This is a test of the direct pairings send via Brevo.</p>' +
    '<p><strong>Foursome 1 - 11:00 AM</strong></p>' +
    '<p style="margin:0;padding-left:20px;">Test Player (Walk)</p>'
  );
  Logger.log(JSON.stringify(result));
}

// ── Broadcast: general email to all active players (not tied to an event) ──

// Wrap the organizer's free-text body in a simple branded email shell.
function buildBroadcastEmailHtml_(bodyText, imageUrl) {
  var safe = (bodyText || '').replace(/\r\n/g, '\n').replace(/\n/g, '<br>');
  var imageSection = imageUrl
    ? '<div style="margin-top:16px;text-align:center;"><img src="' + imageUrl + '" alt="" style="max-width:100%;height:auto;border-radius:8px;"></div>'
    : '';
  return '<meta charset="utf-8">' +
    '<div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;max-width:600px;margin:0 auto;padding:20px;">' +
    '<div style="background:#1a5276;color:white;padding:20px;border-radius:12px 12px 0 0;text-align:center;">' +
    '<div style="font-size:32px;">&#9971;</div>' +
    '<h1 style="margin:8px 0;font-size:22px;">RonsTeeBallers</h1>' +
    '</div>' +
    '<div style="background:white;padding:24px;border:1px solid #e0e8f0;border-top:none;font-size:16px;color:#1a2332;line-height:1.6;">' +
    safe + imageSection +
    '</div>' +
    '<div style="background:#f0f4f8;padding:12px;border-radius:0 0 12px 12px;text-align:center;">' +
    '<p style="color:#aab7c4;font-size:12px;margin:0;">RonsTeeBallers Golf Group</p>' +
    '</div>' +
    '</div>';
}

// Count active players with a valid email (no PII returned). Unguarded — safe to preview.
// Single source of truth for who receives an email for a given group.
// group: 'Main Group' | 'Indian Lakes' | 'X-Golf' | 'Indian Lakes & X-Golf'.
// Rules: a player must be Active (col E = 'Yes') AND have a valid email to be included
// at all - inactive players are NEVER emailed. Indian Lakes (col F) / X-Golf (col G)
// groups additionally require that group's own flag = 'Yes'. Optional excludeResponded
// is a map {lowercased "Last, First": true} of people to drop (used for reminders).
// Returns [{lastFirst, firstName, lastName, email, slug}, ...] in sheet order.
function selectRecipients_(group, excludeResponded) {
  var playersData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players').getDataRange().getValues();
  var out = [];
  for (var p = 1; p < playersData.length; p++) {
    var active = playersData[p][4].toString().trim();
    if (active !== 'Yes') continue;                       // never email inactive players
    var email = playersData[p][3].toString().trim();
    if (!email || email.indexOf('@') === -1) continue;
    var indianLakes = playersData[p][5].toString().trim();  // col F Yes/No
    var xGolf = playersData[p][6].toString().trim();         // col G Yes/No

    var include = false;
    if (group === 'Main Group') include = true;
    else if (group === 'Indian Lakes') include = (indianLakes === 'Yes');
    else if (group === 'X-Golf') include = (xGolf === 'Yes');
    else if (group === 'Indian Lakes & X-Golf') include = (indianLakes === 'Yes' || xGolf === 'Yes');
    if (!include) continue;

    var lastFirst = playersData[p][0].toString().trim();
    if (excludeResponded && excludeResponded[lastFirst.toLowerCase()]) continue;

    out.push({
      lastFirst: lastFirst,
      firstName: playersData[p][1].toString().trim(),
      lastName: playersData[p][2].toString().trim(),
      email: email,
      slug: lastFirst.toLowerCase()
        .replace(/,\s*/g, '-').replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '')
    });
  }
  return out;
}

function broadcastPreview(params) {
  try {
    var group = (params && params.group) || 'Main Group';
    var count = selectRecipients_(group).length;
    // If the organizer has typed a message, also return the exact email HTML
    // so the broadcast modal can show a preview (mirrors previewInvite).
    var body = (params && params.body) ? params.body.toString() : '';
    var html = body.trim() ? buildBroadcastEmailHtml_(body, (params.imageUrl || '').toString()) : '';
    return ContentService.createTextOutput(JSON.stringify({ success: true, count: count, group: group, html: html }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Send a general email (custom subject + body) to active players in the chosen group
// (Main Group / Indian Lakes / X-Golf). Recipient selection via selectRecipients_.
function broadcastEmail(params) {
  try {
    var subject = (params.subject || '').toString().trim();
    var body = (params.body || '').toString();
    if (!subject) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Subject is empty'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (!body.trim()) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Message body is empty'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var group = (params.group || 'Main Group').toString();
    var recipients = selectRecipients_(group);
    if (recipients.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({error: 'No active players with a valid email for group: ' + group}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var htmlBody = buildBroadcastEmailHtml_(body, (params.imageUrl || '').toString());
    var sent = 0;
    var errors = [];
    recipients.forEach(function(r) {
      var res = sendBrevoEmail_(r.email, r.firstName, subject, htmlBody);
      if (res.ok) sent++; else errors.push(r.email + ': ' + res.error);
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true, sent: sent, total: recipients.length, errors: errors
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Commit a pasted image to the GitHub Pages repo (images/ folder) so it can be
// embedded in broadcast emails by URL. Expects { filename, content } where content
// is base64 (no "data:" prefix). Requires Script Property GITHUB_TOKEN — a
// fine-grained PAT with Contents: read/write on RonsTeeBallers/ronsteeballers.github.io.
function uploadImage(data) {
  try {
    var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    if (!token) {
      return ContentService.createTextOutput(JSON.stringify({error: 'GITHUB_TOKEN script property not set'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var filename = (data.filename || '').toString().replace(/[^a-zA-Z0-9._-]/g, '');
    var content = (data.content || '').toString();
    if (!filename || !content) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Missing filename or content'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var path = 'images/' + filename;
    var apiUrl = 'https://api.github.com/repos/RonsTeeBallers/ronsteeballers.github.io/contents/' + path;
    var resp = UrlFetchApp.fetch(apiUrl, {
      method: 'put',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'RonsTeeBallers-AppsScript'
      },
      payload: JSON.stringify({
        message: 'Add broadcast image ' + filename,
        content: content,
        branch: 'main'
      }),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 200 || code === 201) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        url: 'https://ronsteeballers.github.io/' + path
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({
      error: 'GitHub upload failed: HTTP ' + code + ' ' + resp.getContentText()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Organizer statistics, computed from Pairing History + Players + Form Responses.
function getStats() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var historySheet = ss.getSheetByName('Pairing History');
    var playersSheet = ss.getSheetByName('Players');
    var formSheet = ss.getSheetByName('Form Responses');

    var coursesSheet = ss.getSheetByName('Courses');

    // ---------- Canonical player names (First Last) + active flags ----------
    var playersData = playersSheet ? playersSheet.getDataRange().getValues() : [];
    var playerCanon = {};     // 'first last' (lower) -> display name
    var activePlayers = [];   // display names of active players
    for (var p = 1; p < playersData.length; p++) {
      var first = (playersData[p][1] || '').toString().trim();
      var last = (playersData[p][2] || '').toString().trim();
      if (!first && !last) continue;
      var disp = (first + ' ' + last).trim();
      playerCanon[disp.toLowerCase()] = disp;
      if ((playersData[p][4] || '').toString().trim().toLowerCase() === 'yes') activePlayers.push(disp);
    }

    // ---------- Course names (to classify history cells that aren't players) ----------
    var coursesData = coursesSheet ? coursesSheet.getDataRange().getValues() : [];
    var courseList = [];
    for (var ci = 1; ci < coursesData.length; ci++) {
      var cn = (coursesData[ci][0] || '').toString().trim();
      if (cn) courseList.push(cn);
    }
    function canonCourse(val) {
      var low = val.toLowerCase();
      for (var z = 0; z < courseList.length; z++) {
        if (low.indexOf(courseList[z].toLowerCase()) !== -1) return courseList[z];
      }
      return val;
    }

    // ---------- Pairing History, aggregated per OUTING (date) so re-pairings don't double count ----------
    var histData = historySheet ? historySheet.getDataRange().getValues() : [];
    var hStart = 0;
    if (histData.length > 0) {
      var hHead = histData[0].join(' ').toLowerCase();
      if (hHead.indexOf('player') !== -1 || hHead.indexOf('date') !== -1 || hHead.indexOf('course') !== -1) hStart = 1;
    }

    var datePlayers = {};  // date -> { playerKey: display }   (only recognized players)
    var datePairs = {};    // date -> { 'ka||kb': true }
    var dateCourse = {};   // date -> canonical course
    var allDates = {};

    for (var r = hStart; r < histData.length; r++) {
      var hrow = histData[r];
      if (!hrow) continue;
      var dateStr = (hrow[0] || '').toString().trim();
      if (!dateStr) continue;

      var names = [];
      var courseMatch = '';
      var lastUnknown = '';
      for (var c = 1; c < hrow.length; c++) {
        var val = (hrow[c] || '').toString().trim();
        if (!val) continue;
        var canonName = playerCanon[val.toLowerCase()];
        if (canonName) {
          names.push(canonName);
        } else {
          var cc = canonCourse(val);
          if (cc !== val || courseList.indexOf(val) !== -1) courseMatch = cc;
          else lastUnknown = val;
        }
      }

      if (names.length > 0) {
        if (!datePlayers[dateStr]) datePlayers[dateStr] = {};
        if (!datePairs[dateStr]) datePairs[dateStr] = {};
        names.forEach(function(n) { datePlayers[dateStr][n.toLowerCase()] = n; });
        for (var a = 0; a < names.length; a++) {
          for (var b = a + 1; b < names.length; b++) {
            var ka = names[a].toLowerCase(), kb = names[b].toLowerCase();
            datePairs[dateStr][(ka < kb ? ka + '||' + kb : kb + '||' + ka)] = true;
          }
        }
        allDates[dateStr] = true;
      }
      var courseForRow = courseMatch || lastUnknown;
      if (courseForRow) { dateCourse[dateStr] = courseForRow; allDates[dateStr] = true; }
    }

    // Rounds played = number of distinct outings a player appears in
    var roundsByKey = {};
    var displayByKey = {};
    var totalRounds = 0;
    Object.keys(datePlayers).forEach(function(d) {
      Object.keys(datePlayers[d]).forEach(function(k) {
        roundsByKey[k] = (roundsByKey[k] || 0) + 1;
        displayByKey[k] = datePlayers[d][k];
        totalRounds++;
      });
    });

    var leaderboard = Object.keys(roundsByKey).map(function(k) {
      return { name: displayByKey[k], rounds: roundsByKey[k] };
    }).sort(function(x, y) { return y.rounds - x.rounds || x.name.localeCompare(y.name); });

    // Partners = distinct outings two players shared a foursome
    var partnerCounts = {};
    Object.keys(datePairs).forEach(function(d) {
      Object.keys(datePairs[d]).forEach(function(pairKey) {
        var parts = pairKey.split('||');
        var ka = parts[0], kb = parts[1];
        if (!partnerCounts[ka]) partnerCounts[ka] = {};
        if (!partnerCounts[kb]) partnerCounts[kb] = {};
        partnerCounts[ka][kb] = (partnerCounts[ka][kb] || 0) + 1;
        partnerCounts[kb][ka] = (partnerCounts[kb][ka] || 0) + 1;
      });
    });
    var topPartners = Object.keys(partnerCounts).map(function(k) {
      var partners = Object.keys(partnerCounts[k]).map(function(pk) {
        return { name: displayByKey[pk] || pk, count: partnerCounts[k][pk] };
      }).sort(function(x, y) { return y.count - x.count || x.name.localeCompare(y.name); }).slice(0, 5);
      return { name: displayByKey[k], partners: partners };
    }).sort(function(x, y) { return x.name.localeCompare(y.name); });

    // Courses = number of distinct outings planned at each course (players not counted)
    var courseOutingCount = {};
    Object.keys(dateCourse).forEach(function(d) {
      var nm = dateCourse[d];
      courseOutingCount[nm] = (courseOutingCount[nm] || 0) + 1;
    });
    var courses = Object.keys(courseOutingCount).map(function(nm) {
      return { course: nm, outings: courseOutingCount[nm] };
    }).sort(function(x, y) { return y.outings - x.outings || x.course.localeCompare(y.course); });

    var outingDates = allDates;
    var activeCount = activePlayers.length;

    // Active players who have never played
    var neverPlayed = activePlayers.filter(function(disp) {
      return !roundsByKey[disp.toLowerCase()];
    }).sort(function(x, y) { return x.localeCompare(y); });

    // ---------- Walk vs Ride from Form Responses (Yes only, last response per event) ----------
    var formData = formSheet ? formSheet.getDataRange().getValues() : [];
    var fStart = 0;
    if (formData.length > 0) {
      var fHead = formData[0].join(' ').toLowerCase();
      if (fHead.indexOf('timestamp') !== -1 || fHead.indexOf('name') !== -1 || fHead.indexOf('walk') !== -1) fStart = 1;
    }
    // Dedup by Timestamp per (player, event) BEFORE filtering to Yes, so a
    // Yes changed to a No drops out and sheet sort order never matters.
    var wrLast = {};   // (nameKey|eventId) -> { ts, playing, wr } (latest by Timestamp)
    var wrName = {};   // nameKey -> display
    for (var f = fStart; f < formData.length; f++) {
      var frow = formData[f];
      var rawName = (frow[1] || '').toString().trim();
      var playing = (frow[2] || '').toString().trim().toLowerCase();
      var wrVal = (frow[3] || '').toString().trim();
      var ev = (frow[6] || '').toString().trim();
      if (rawName === '') continue;
      var disp2 = rawName;
      if (rawName.indexOf(',') !== -1) {
        var parts2 = rawName.split(',');
        disp2 = ((parts2[1] || '').trim() + ' ' + (parts2[0] || '').trim()).trim();
      }
      var nk = disp2.toLowerCase();
      wrName[nk] = disp2;
      var wrKey = nk + '|' + ev;
      var fts = (frow[0] instanceof Date) ? frow[0].getTime() : 0;
      if (!wrLast[wrKey] || fts >= wrLast[wrKey].ts) {
        wrLast[wrKey] = { ts: fts, playing: playing, wr: wrVal.toLowerCase() };
      }
    }
    var wrTally = {};
    Object.keys(wrLast).forEach(function(key) {
      if (wrLast[key].playing !== 'yes') return;
      var nk = key.substring(0, key.lastIndexOf('|'));
      if (!wrTally[nk]) wrTally[nk] = { walk: 0, ride: 0, either: 0 };
      var v = wrLast[key].wr;
      if (v === 'walk') wrTally[nk].walk++;
      else if (v === 'ride') wrTally[nk].ride++;
      else wrTally[nk].either++;
    });
    var walkRide = Object.keys(wrTally).map(function(nk) {
      var t = wrTally[nk];
      return { name: wrName[nk], walk: t.walk, ride: t.ride, either: t.either, total: t.walk + t.ride + t.either };
    }).sort(function(x, y) { return y.total - x.total || x.name.localeCompare(y.name); });

    var result = {
      totalOutings: Object.keys(outingDates).length,
      totalRounds: totalRounds,
      activePlayers: activeCount,
      neverPlayedCount: neverPlayed.length,
      leaderboard: leaderboard,
      topPartners: topPartners,
      neverPlayed: neverPlayed,
      walkRide: walkRide,
      courses: courses
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getEventInfo(eventId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var coursesSheet = ss.getSheetByName('Courses');

    var found = findEventRow_(eventId);
    if (!found) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var venueName = found.row[1].toString();
    var startTime = found.row[2];
    var timeStr = startTime instanceof Date
      ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'h:mm a')
      : startTime.toString();

    var teeTimeInterval = 9;
    var courseData = coursesSheet.getDataRange().getValues();
    for (var c = 1; c < courseData.length; c++) {
      if (courseData[c][0].toString() === venueName) {
        var interval = parseInt(courseData[c][10]);
        if (!isNaN(interval) && interval > 0) teeTimeInterval = interval;
        break;
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      event: {
        venue: venueName,
        time: timeStr,
        interval: teeTimeInterval
      }
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// Builds the exact invite email HTML. Used by both the live send and the preview
// so the organizer's preview is identical to what recipients receive.
function buildInviteEmailHtml_(o) {
  var feeSection = '';
  if (o.total) {
    feeSection = '<p style="color:#5d6d7e;font-size:14px;">Green Fee: ' + o.greenFee +
      ' &nbsp;|&nbsp; Cart: ' + o.cartFee +
      ' &nbsp;|&nbsp; <strong>Total: ' + o.total + '</strong></p>';
  }
  var venueLink = o.venueUrl
    ? '<a href="' + o.venueUrl + '" style="color:#1a5276;">' + o.venueName + '</a>'
    : o.venueName;
  var commentSection = o.comment
    ? '<p style="background:#eaf2ff;padding:12px;border-radius:8px;border-left:3px solid #1a5276;">' + o.comment + '</p>'
    : '';
  // Event Notes are the organizer's free text from the initial invitation - req'd to be
  // stripped from REMIND emails so the reminder only carries the current message + roster.
  var notesSection = (o.notes && !o.isReminder)
    ? '<p style="color:#5d6d7e;font-size:15px;line-height:1.5;margin:0 0 12px;">' + o.notes + '</p>'
    : '';
  var courseNotesSection = o.courseNotes
    ? '<p style="color:#5d6d7e;font-size:14px;line-height:1.5;margin:0 0 12px;">' + o.courseNotes + '</p>'
    : '';
  var signedUpSection = '';
  if (o.isReminder && o.signedUpNames && o.signedUpNames.length) {
    signedUpSection = '<div style="margin:0 0 16px;">' +
      '<p style="font-weight:700;color:#1a2332;margin:0 0 6px;">Players Signed Up</p>' +
      '<p style="color:#1a2332;margin:0;line-height:1.6;">' +
      o.signedUpNames.map(escapeHtml_).join('<br>') +
      '</p></div>';
  }
  var detailsBlock = '<div style="background:#f0f4f8;padding:16px;border-radius:8px;margin:16px 0;">' +
    '<p style="margin:4px 0;font-size:16px;font-weight:700;color:#1a2332;">&#128197; ' + o.formattedDate + '</p>' +
    '<p style="margin:4px 0;font-size:16px;color:#1a2332;">&#128205; ' + venueLink + '</p>' +
    '<p style="margin:4px 0;font-size:16px;color:#1a2332;">&#9200; ' + o.timeStr + ' (first tee time)</p>' +
    '<p style="margin:4px 0;font-size:14px;color:#5d6d7e;">' + o.slotsReserved + ' tee times reserved under the name Ron Blanton</p>' +
    '</div>';
  var buttonsBlock = '<div style="text-align:center;margin:24px 0;">' +
    '<a href="' + o.rsvpUrl + '" style="display:inline-block;background:#1a5276;color:white;padding:16px 32px;border-radius:10px;text-decoration:none;font-size:18px;font-weight:700;margin:0 8px;">&#9989; I\'M IN</a>' +
    '<a href="' + o.rsvpUrl + '" style="display:inline-block;background:#922b21;color:white;padding:16px 32px;border-radius:10px;text-decoration:none;font-size:18px;font-weight:700;margin:0 8px;">&#10060; I\'M OUT</a>' +
    '</div>';
  var signupLine = '<p style="text-align:center;color:#5d6d7e;font-size:14px;">See who else is playing: <a href="' + o.signupUrl + '" style="color:#1a5276;">View Signup List</a></p>';

  var bodyContent;
  if (o.isReminder) {
    // Organizer's message is the reason for the reminder - sits right under the greeting
    // (one blank line between), followed by the roster, then the invitation itself so the
    // recipient never has to hunt for the original email or ask to be added.
    bodyContent =
      '<p style="font-size:18px;font-weight:700;color:#1a2332;margin:0 0 16px;">Hi ' + o.firstName + '!</p>' +
      commentSection +
      signedUpSection +
      detailsBlock +
      feeSection +
      courseNotesSection +
      buttonsBlock +
      signupLine;
  } else {
    bodyContent =
      '<p style="font-size:18px;font-weight:700;color:#1a2332;">Hi ' + o.firstName + '!</p>' +
      '<p style="color:#5d6d7e;">You\'re invited to join us for golf:</p>' +
      detailsBlock +
      notesSection +
      feeSection +
      courseNotesSection +
      commentSection +
      buttonsBlock +
      signupLine;
  }

  return '<meta charset="utf-8">' +
    '<div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;max-width:600px;margin:0 auto;padding:20px;">' +
    '<div style="background:#1a5276;color:white;padding:20px;border-radius:12px 12px 0 0;text-align:center;">' +
    '<div style="font-size:32px;">&#9971;</div>' +
    '<h1 style="margin:8px 0;font-size:22px;">' + (o.isReminder ? 'Golf Outing Reminder' : 'Golf Outing Invitation') + '</h1>' +
    '</div>' +
    '<div style="background:white;padding:24px;border:1px solid #e0e8f0;border-top:none;">' +
    bodyContent +
    '</div>' +
    '<div style="background:#f0f4f8;padding:12px;border-radius:0 0 12px 12px;text-align:center;">' +
    '<p style="color:#aab7c4;font-size:12px;margin:0;">See you on the course!</p>' +
    '</div>' +
    '</div>';
}

// Minimal HTML-escaping for user-entered text (e.g. guest names typed into the public
// RSVP form) that gets embedded directly into email HTML.
function escapeHtml_(s) {
  return (s || '').toString()
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Returns a map {lowercased "Last, First": 'Yes'|'No'} of the latest RSVP response
// per player for the given event - used to filter REMIND recipients by audience.
function getResponseStatusMap_(eventId) {
  var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  var formData = formSheet.getDataRange().getValues();
  var latest = latestRowsByPlayer_(formData, eventId);
  var status = {};
  for (var name in latest) {
    var srow = latest[name].row;
    status[name.toLowerCase()] = srow[2] ? srow[2].toString().trim() : '';
  }
  return status;
}

// Builds the exclude-set for a REMIND send from the organizer's chosen audience:
// 'all' = no exclusions, 'unresponded_or_no' = skip only players who said Yes,
// 'unresponded' (default) = skip anyone who answered at all (Yes or No).
function computeRemindExcludeSet_(statusMap, audience) {
  var exclude = {};
  if (audience === 'all') return exclude;
  for (var nm in statusMap) {
    if (audience === 'unresponded_or_no') {
      if (statusMap[nm] === 'Yes') exclude[nm] = true;
    } else {
      exclude[nm] = true;
    }
  }
  return exclude;
}

// Returns display names ("First Last") of everyone with a Yes response for eventId,
// including any guests, sorted alphabetically. Used by the "Players Signed Up" list
// in REMIND emails.
function getSignedUpNames_(eventId) {
  var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  var formData = formSheet.getDataRange().getValues();
  var latest = latestRowsByPlayer_(formData, eventId);

  var names = [];
  for (var lastFirst in latest) {
    var nrow = latest[lastFirst].row;
    if ((nrow[2] ? nrow[2].toString().trim() : '') !== 'Yes') continue;
    var parts = lastFirst.split(',');
    names.push(parts.length > 1 ? parts[1].trim() + ' ' + parts[0].trim() : lastFirst);
    var guest = nrow[7] ? nrow[7].toString().trim() : '';
    if (guest) names.push(guest);
  }
  names.sort(function(a, b) { return a.localeCompare(b); });
  return names;
}

// Build a one-recipient preview of the invite (no email is sent).
function previewInvite(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var coursesSheet = ss.getSheetByName('Courses');

    var eventId = params.eventId;
    var mailingList = params.mailingList;
    var comment = params.comment || '';
    var remindOnly = (params.remindOnly === 'true' || params.remindOnly === true);
    var audience = params.audience || 'unresponded';
    var responded = remindOnly ? computeRemindExcludeSet_(getResponseStatusMap_(eventId), audience) : {};

    var found = findEventRow_(eventId);
    if (!found) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var eventRow = found.row;

    var venueName = eventRow[1].toString();
    var startTime = eventRow[2];
    var timeStr = startTime instanceof Date
      ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'h:mm a')
      : startTime.toString();
    var slotsReserved = eventRow[3].toString();
    var formattedDate = eventRow[0] instanceof Date
      ? Utilities.formatDate(eventRow[0], Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')
      : eventRow[0].toString();

    var coursesData = coursesSheet.getDataRange().getValues();
    var greenFee = '', cartFee = '', total = '', venueUrl = '', courseNotes = '';
    for (var c = 1; c < coursesData.length; c++) {
      if (coursesData[c][0].toString() === venueName) {
        var green = parseFloat(coursesData[c][7]) || 0;
        var cart = parseFloat(coursesData[c][8]) || 0;
        greenFee = green > 0 ? '$' + green.toFixed(2) : '';
        cartFee = cart > 0 ? '$' + cart.toFixed(2) : '';
        total = (green + cart) > 0 ? '$' + (green + cart).toFixed(2) : '';
        venueUrl = coursesData[c][6] ? coursesData[c][6].toString() : '';
        courseNotes = coursesData[c][11] ? coursesData[c][11].toString() : '';  // col L Course Notes
        break;
      }
    }

    // Count recipients + pick a sample name (same active-gated selection as the send).
    var recipients = selectRecipients_(mailingList, responded);
    var count = recipients.length;
    var sampleName = count ? (recipients[0].firstName || 'Friend') : 'Friend';
    var sampleSlug = count ? recipients[0].slug : 'sample-player';

    var baseUrl = 'https://ronsteeballers.github.io';
    var html = buildInviteEmailHtml_({
      firstName: sampleName,
      formattedDate: formattedDate, venueName: venueName, venueUrl: venueUrl,
      timeStr: timeStr, slotsReserved: slotsReserved,
      greenFee: greenFee, cartFee: cartFee, total: total,
      comment: comment,
      isReminder: remindOnly,
      notes: eventRow[6] ? eventRow[6].toString() : '',
      courseNotes: courseNotes,
      signedUpNames: remindOnly ? getSignedUpNames_(eventId) : [],
      rsvpUrl: baseUrl + '/rsvp.html?event=' + eventId + '&player=' + sampleSlug,
      signupUrl: baseUrl + '/signup.html?event=' + eventId
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      html: html,
      subject: (remindOnly ? 'Reminder: ' : '') + 'Golf Outing - ' + formattedDate + ' at ' + venueName,
      count: count,
      mailingList: mailingList,
      sampleName: sampleName
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendInviteEmails(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var coursesSheet = ss.getSheetByName('Courses');

    var eventId = params.eventId;
    var mailingList = params.mailingList;
    var comment = params.comment || '';
    var remindOnly = (params.remindOnly === 'true' || params.remindOnly === true);
    var audience = params.audience || 'unresponded';
    var responded = remindOnly ? computeRemindExcludeSet_(getResponseStatusMap_(eventId), audience) : {};

    // Get event details
    var found = findEventRow_(eventId);
    if (!found) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var eventRow = found.row;

    var venueName = eventRow[1].toString();
    var startTime = eventRow[2];
    var timeStr = startTime instanceof Date
      ? Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'h:mm a')
      : startTime.toString();
    var slotsReserved = eventRow[3].toString();
    var formattedDate = eventRow[0] instanceof Date
      ? Utilities.formatDate(eventRow[0], Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')
      : eventRow[0].toString();

    // Get venue details
    var coursesData = coursesSheet.getDataRange().getValues();
    var greenFee = '';
    var cartFee = '';
    var total = '';
    var venueUrl = '';
    var courseNotes = '';
    for (var c = 1; c < coursesData.length; c++) {
      if (coursesData[c][0].toString() === venueName) {
        var green = parseFloat(coursesData[c][7]) || 0;
        var cart = parseFloat(coursesData[c][8]) || 0;
        greenFee = green > 0 ? '$' + green.toFixed(2) : '';
        cartFee = cart > 0 ? '$' + cart.toFixed(2) : '';
        total = (green + cart) > 0 ? '$' + (green + cart).toFixed(2) : '';
        venueUrl = coursesData[c][6] ? coursesData[c][6].toString() : '';
        courseNotes = coursesData[c][11] ? coursesData[c][11].toString() : '';  // col L Course Notes
        break;
      }
    }

    // Get players for this mailing list (active-gated; reminders skip responders).
    var recipients = selectRecipients_(mailingList, responded);

    if (recipients.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({error: 'No recipients found for mailing list: ' + mailingList}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var baseUrl = 'https://ronsteeballers.github.io';
    var signupUrl = baseUrl + '/signup.html?event=' + eventId;
    var sent = 0;
    var errors = [];
    var signedUpNames = remindOnly ? getSignedUpNames_(eventId) : [];

    // Send personalized email to each recipient
    recipients.forEach(function(recipient) {
      var rsvpUrl = baseUrl + '/rsvp.html?event=' + eventId + '&player=' + recipient.slug;

      var htmlBody = buildInviteEmailHtml_({
        firstName: recipient.firstName,
        formattedDate: formattedDate, venueName: venueName, venueUrl: venueUrl,
        timeStr: timeStr, slotsReserved: slotsReserved,
        greenFee: greenFee, cartFee: cartFee, total: total,
        comment: comment, isReminder: remindOnly, notes: eventRow[6] ? eventRow[6].toString() : '',
        courseNotes: courseNotes,
        signedUpNames: signedUpNames,
        rsvpUrl: rsvpUrl, signupUrl: signupUrl
      });

      var subject = (remindOnly ? 'Reminder: ' : '') + 'Golf Outing - ' + formattedDate + ' at ' + venueName;
      var sendResult = sendBrevoEmail_(recipient.email, recipient.firstName, subject, htmlBody);
      if (sendResult.ok) {
        sent++;
      } else {
        errors.push(recipient.email + ': ' + sendResult.error);
      }
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      sent: sent,
      errors: errors,
      total: recipients.length
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// Public RSVP endpoint (unauthenticated by design - reached from emailed links).
// Validates the submission against the sheets before writing anything: the
// player must exist in the Players tab and the event must exist and be Open.
function submitRSVP(params) {
  var playerName = (params.playerName || '').toString().trim();
  var eventId = (params.eventId || '').toString().trim();

  var found = findEventRow_(eventId);
  if (!found) {
    return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (found.row[5] !== 'Open') {
    return ContentService.createTextOutput(JSON.stringify({
      error: 'This outing is closed - pairings have already been sent. Contact Ron if your plans changed.',
      closed: true
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var playersData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players').getDataRange().getValues();
  var canonical = '';
  for (var i = 1; i < playersData.length; i++) {
    var lastFirst = playersData[i][0].toString().trim();
    if (lastFirst && lastFirst.toLowerCase() === playerName.toLowerCase()) {
      canonical = lastFirst;   // write the roster's exact casing
      break;
    }
  }
  if (!canonical) {
    return ContentService.createTextOutput(JSON.stringify({error: 'Player not found - please use the link from your invite email.'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return writeRsvpRow_(canonical, params);
}

// Shared by submitRSVP (player-submitted, playerName trusted from the RSVP link)
// and organizerSubmitRSVP (organizer-submitted, playerName resolved server-side
// from the Players tab). Writes one Form Responses row.
function writeRsvpRow_(playerName, params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName('Form Responses');

    // Strip the pairing-payload delimiters (| ; ,) from guest names - a guest
    // named with one of them would corrupt the foursomes string sent by the
    // pairing page.
    var guestFirst = (params.guestFirst || '').toString().replace(/[|;,]/g, ' ').trim();
    var guestLast = (params.guestLast || '').toString().replace(/[|;,]/g, ' ').trim();
    var guestName = (guestFirst + ' ' + guestLast).replace(/\s+/g, ' ').trim();
    var playingYes = (params.playing === 'Yes');

    // A guest must be a true guest, not an existing active player
    if (playingYes && guestName) {
      var playersData = ss.getSheetByName('Players').getDataRange().getValues();
      var guestKey = guestName.toLowerCase();
      for (var i = 1; i < playersData.length; i++) {
        if ((playersData[i][4] || '').toString().trim().toLowerCase() !== 'yes') continue;
        var full = ((playersData[i][1] || '') + ' ' + (playersData[i][2] || ''))
          .toString().toLowerCase().replace(/\s+/g, ' ').trim();
        if (full && full === guestKey) {
          return ContentService.createTextOutput(JSON.stringify({
            error: '"' + guestName + '" is already an active player. Please have them RSVP with their own link, or enter a different guest.'
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    var timestamp = new Date();
    var row = [
      timestamp,
      playerName || '',
      params.playing || '',
      params.walkRide || '',
      params.scoring || '',
      params.comments || '',
      params.eventId || '',
      playingYes ? guestName : ''
    ];

    var nextRow = formSheet.getLastRow() + 1;
    var numCols = row.length;
    formSheet.getRange(nextRow, 1, 1, numCols).setValues([row]);
    formSheet.getRange(nextRow, 7).setNumberFormat('@STRING@');
    formSheet.getRange(nextRow, 7).setValue(params.eventId || '');

    return ContentService.createTextOutput(JSON.stringify({success: true}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Look up a player in the Players tab by their RSVP-link slug (same slug algorithm
// as getEventData). Returns { lastFirst, first, last } or null.
function resolvePlayerBySlug_(slug) {
  var playersData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players').getDataRange().getValues();
  var target = (slug || '').toString().toLowerCase();
  for (var i = 1; i < playersData.length; i++) {
    var lastFirst = playersData[i][0].toString().trim();
    var s = lastFirst.toLowerCase().replace(/,\s*/g, '-').replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
    if (s === target) {
      return {
        lastFirst: lastFirst,
        first: playersData[i][1].toString().trim(),
        last: playersData[i][2].toString().trim()
      };
    }
  }
  return null;
}

// Full player roster for the organizer's "Respond as a player" tool. Guarded by
// passcode since it lists every player (active and inactive) for impersonation.
function getPlayers() {
  try {
    var playersData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players').getDataRange().getValues();
    var players = [];
    for (var i = 1; i < playersData.length; i++) {
      var lastFirst = playersData[i][0].toString().trim();
      if (!lastFirst) continue;
      var first = playersData[i][1].toString().trim();
      var last = playersData[i][2].toString().trim();
      var active = (playersData[i][4] ? playersData[i][4].toString().trim() : '') === 'Yes';
      var slug = lastFirst.toLowerCase().replace(/,\s*/g, '-').replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
      players.push({ name: (first + ' ' + last).trim(), slug: slug, active: active });
    }
    players.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return ContentService.createTextOutput(JSON.stringify({players: players}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Organizer-submitted RSVP on behalf of a player (e.g. they called in their response).
// Resolves the player server-side from playerSlug so the written name always matches
// the canonical Players tab entry, then writes the same row shape as submitRSVP.
function organizerSubmitRSVP(params) {
  var slug = (params.playerSlug || '').toString().trim();
  var eventId = (params.eventId || '').toString().trim();
  if (!slug || !eventId) {
    return ContentService.createTextOutput(JSON.stringify({error: 'Missing player or event'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var player = resolvePlayerBySlug_(slug);
  if (!player) {
    return ContentService.createTextOutput(JSON.stringify({error: 'Player not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return writeRsvpRow_(player.lastFirst, params);
}

// Same slug algorithm as resolvePlayerBySlug_/getPlayers, factored out for the
// player add/edit tool (needs it for both lookup and duplicate-name detection).
function slugify_(lastFirst) {
  return (lastFirst || '').toString().toLowerCase()
    .replace(/,\s*/g, '-').replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
}

// Full detail for one player (used to populate the edit form). Guarded — returns
// email/phone, which getPlayers deliberately omits.
function getPlayerDetail(params) {
  try {
    var slug = (params.slug || '').toString().trim().toLowerCase();
    var playersData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players').getDataRange().getValues();
    for (var i = 1; i < playersData.length; i++) {
      var lastFirst = playersData[i][0].toString().trim();
      if (slugify_(lastFirst) !== slug) continue;
      return ContentService.createTextOutput(JSON.stringify({
        player: {
          slug: slug,
          first: playersData[i][1] ? playersData[i][1].toString().trim() : '',
          last: playersData[i][2] ? playersData[i][2].toString().trim() : '',
          email: playersData[i][3] ? playersData[i][3].toString().trim() : '',
          active: (playersData[i][4] ? playersData[i][4].toString().trim() : '') === 'Yes',
          indianLakes: (playersData[i][5] ? playersData[i][5].toString().trim() : '') === 'Yes',
          xGolf: (playersData[i][6] ? playersData[i][6].toString().trim() : '') === 'Yes',
          phone: playersData[i][7] ? playersData[i][7].toString().trim() : ''
        }
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({error: 'Player not found'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Add a new player, or edit an existing one (originalSlug identifies which row to
// overwrite; empty originalSlug means "add new"). Column H (Phone) is new here —
// columns A-G are unchanged, see players-tab-schema.
function savePlayer(params) {
  try {
    var first = (params.first || '').toString().trim();
    var last = (params.last || '').toString().trim();
    if (!first || !last) {
      return ContentService.createTextOutput(JSON.stringify({error: 'First and last name are required'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var email = (params.email || '').toString().trim();
    var phone = (params.phone || '').toString().trim();
    var active = (params.active === 'Yes') ? 'Yes' : 'No';
    var indianLakes = (params.indianLakes === 'Yes') ? 'Yes' : 'No';
    var xGolf = (params.xGolf === 'Yes') ? 'Yes' : 'No';
    var lastFirst = last + ', ' + first;
    var newSlug = slugify_(lastFirst);
    var originalSlug = (params.originalSlug || '').toString().trim().toLowerCase();

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players');
    var playersData = sheet.getDataRange().getValues();

    // Block collisions with a DIFFERENT existing player landing on the same slug
    // (slug identifies players across RSVP links and this tool - must stay unique).
    for (var i = 1; i < playersData.length; i++) {
      var rowSlug = slugify_(playersData[i][0].toString().trim());
      if (rowSlug === newSlug && rowSlug !== originalSlug) {
        return ContentService.createTextOutput(JSON.stringify({
          error: 'A player named "' + lastFirst + '" already exists.'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    var rowValues = [lastFirst, first, last, email, active, indianLakes, xGolf, phone];

    if (originalSlug) {
      var rowIndex = -1;
      for (var j = 1; j < playersData.length; j++) {
        if (slugify_(playersData[j][0].toString().trim()) === originalSlug) { rowIndex = j + 1; break; }
      }
      if (rowIndex === -1) {
        return ContentService.createTextOutput(JSON.stringify({error: 'Player not found (may have been edited elsewhere) - reopen and try again.'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return ContentService.createTextOutput(JSON.stringify({success: true, slug: newSlug}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------------------------------------------------------------------------
// Tee-time availability finder (Phase 1: read-only reporter, no booking).
// Given a date, earliest start time, and a required number of consecutive
// open tee times, checks each course's live tee sheet (Courses col M =
// Online Booking URL) and reports which ones can host the group.
// ---------------------------------------------------------------------------

function findTeeTimeAvailability(params) {
  try {
    var date = (params.date || '').toString().trim();          // yyyy-MM-dd
    var earliestTime = (params.earliestTime || '00:00').toString().trim(); // HH:MM 24h
    var minConsecutive = parseInt(params.minConsecutive) || 5;
    var playersPerSlot = 4;

    if (!date) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Date is required'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var earliestMinutes = timeStrToMinutes_(earliestTime);
    var coursesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Courses');
    var coursesData = coursesSheet.getDataRange().getValues();

    var results = [];
    for (var i = 1; i < coursesData.length; i++) {
      var row = coursesData[i];
      var courseName = (row[0] || '').toString().trim();
      if (!courseName) continue;

      var bookingUrl = (row[12] || '').toString().trim();   // col M
      var interval = parseInt(row[10]) || 9;                // col K

      if (!bookingUrl) {
        results.push({ course: courseName, status: 'manual' });
        continue;
      }

      var platform = getPlatformInfo_(bookingUrl);
      if (!platform) {
        results.push({ course: courseName, status: 'manual', bookingUrl: bookingUrl });
        continue;
      }

      try {
        var slots;
        if (platform.type === 'foreup') {
          slots = fetchForeUpTimes_(platform.host, platform.scheduleId, formatDateForForeUp_(date));
        } else if (platform.type === 'golfrev') {
          slots = fetchGolfRevTimes_(platform.courseid, platform.htc, formatDateForGolfRev_(date));
        } else if (platform.type === 'chronogolf') {
          slots = fetchChronogolfTimes_(platform.clubId, date, playersPerSlot);
        } else {
          slots = null;
        }

        if (!slots) {
          results.push({ course: courseName, status: 'manual', bookingUrl: bookingUrl });
          continue;
        }

        var window = findConsecutiveWindow_(slots, interval, playersPerSlot, minConsecutive, earliestMinutes);
        if (window) {
          results.push({
            course: courseName,
            status: 'available',
            firstTime: window[0].time,
            count: window.length,
            bookingUrl: bookingUrl
          });
        } else {
          results.push({ course: courseName, status: 'none', bookingUrl: bookingUrl });
        }
      } catch (courseErr) {
        results.push({ course: courseName, status: 'error', bookingUrl: bookingUrl, error: courseErr.message });
      }
    }

    var order = { available: 0, none: 1, manual: 2, error: 3 };
    results.sort(function(a, b) { return order[a.status] - order[b.status]; });

    return ContentService.createTextOutput(JSON.stringify({ results: results }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Detects the booking platform from an Online Booking URL and pulls out the
// IDs each platform's API needs. Returns null for unsupported platforms.
function getPlatformInfo_(bookingUrl) {
  var m = bookingUrl.match(/\/\/([^\/]+)\/.*\/booking\/(\d+)\/(\d+)/);
  if (m) {
    return { type: 'foreup', host: m[1], scheduleId: m[3] };
  }

  if (bookingUrl.indexOf('golfrev.com') !== -1) {
    var q = parseQueryParams_(bookingUrl);
    if (q.htc && q.courseid) {
      return { type: 'golfrev', htc: q.htc, courseid: q.courseid };
    }
  }

  m = bookingUrl.match(/chronogolf\.com\/club\/(\d+)/);
  if (m) {
    return { type: 'chronogolf', clubId: m[1] };
  }

  return null;
}

function parseQueryParams_(url) {
  var params = {};
  var qIndex = url.indexOf('?');
  if (qIndex === -1) return params;
  var qs = url.substring(qIndex + 1).split('#')[0];
  qs.split('&').forEach(function(pair) {
    if (!pair) return;
    var kv = pair.split('=');
    var key = decodeURIComponent(kv[0] || '');
    if (key) params[key] = decodeURIComponent((kv[1] || '').replace(/\+/g, ' '));
  });
  return params;
}

function formatDateForForeUp_(isoDate) {
  var parts = isoDate.split('-');            // yyyy-MM-dd
  return parts[1] + '-' + parts[2] + '-' + parts[0];  // MM-DD-yyyy
}

function formatDateForGolfRev_(isoDate) {
  var parts = isoDate.split('-');
  return parts[1] + '/' + parts[2] + '/' + parts[0];  // MM/dd/yyyy
}

function timeStrToMinutes_(hhmm) {
  var parts = hhmm.split(':');
  var h = parseInt(parts[0]) || 0;
  var m = parseInt(parts[1]) || 0;
  return h * 60 + m;
}

// foreUp public JSON API - returns only slots that still have room.
function fetchForeUpTimes_(host, scheduleId, dateStr) {
  var url = 'https://' + host + '/index.php/api/booking/times?time=all&date=' + encodeURIComponent(dateStr) +
    '&holes=all&players=0&booking_class=&schedule_id=' + scheduleId +
    '&schedule_ids%5B%5D=' + scheduleId + '&specials_only=0&api_key=no_external_api_key';
  var resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'X-Requested-With': 'XMLHttpRequest' }
  });
  var data = JSON.parse(resp.getContentText());
  if (!data || !data.length) return [];

  return data.map(function(slot) {
    var timePart = (slot.time || '').split(' ')[1] || '';   // "yyyy-MM-dd HH:mm" -> "HH:mm"
    return { time: timePart, availableSpots: parseInt(slot.available_spots) || 0 };
  }).filter(function(s) { return s.time; });
}

// GolfRev HTML fragment endpoint - one card per bookable slot.
function fetchGolfRevTimes_(courseid, htc, dateStr) {
  var url = 'https://www.golfrev.com/go/tee_times/teetime_table_html.asp?c=' + courseid +
    '&s=' + encodeURIComponent(dateStr) + '&h=' + htc + '&specials=&reset=yes&snapshot=no';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var html = resp.getContentText();

  var slots = [];
  var re = /showBooking\('([^']+)',(\d+),(\d+),(\d+),(\d+),/g;
  var match;
  while ((match = re.exec(html))) {
    var hour = parseInt(match[3]);
    var minute = parseInt(match[4]);
    var players = parseInt(match[5]);
    var time = (hour < 10 ? '0' + hour : '' + hour) + ':' + (minute < 10 ? '0' + minute : '' + minute);
    slots.push({ time: time, availableSpots: players });
  }
  return slots;
}

// Chronogolf JSON API. Course id + affiliation type id aren't in the booking
// URL, so this scrapes the club page's embedded __NEXT_DATA__ once per call
// to find them, then queries the teetimes endpoint for the requested date.
function fetchChronogolfTimes_(clubId, isoDate, players) {
  var pageResp = UrlFetchApp.fetch('https://www.chronogolf.com/club/' + clubId, { muteHttpExceptions: true });
  var html = pageResp.getContentText();
  var m = html.match(/__NEXT_DATA__[^>]*>([\s\S]*?)<\/script>/);
  if (!m) return [];

  var nextData = JSON.parse(m[1]);
  var club = nextData.props && nextData.props.pageProps && nextData.props.pageProps.club;
  if (!club || !club.courses || !club.courses.length) return [];

  var courseId = club.courses[0].id;
  var affiliationTypeId = club.defaultAffiliationTypeId;

  var url = 'https://www.chronogolf.com/marketplace/clubs/' + clubId + '/teetimes?date=' +
    encodeURIComponent(isoDate) + '&course_id=' + courseId + '&nb_holes=18&nb_players=' + players +
    '&affiliation_type_ids%5B%5D=' + affiliationTypeId;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'Accept': 'application/json' } });
  var data = JSON.parse(resp.getContentText());
  if (!data || !data.length) return [];

  return data
    .filter(function(slot) { return slot.out_of_capacity === false; })
    .map(function(slot) { return { time: slot.start_time, availableSpots: players }; });
}

// Finds the earliest run of `minCount` tee times, each with room for
// `playersNeeded`, spaced exactly `intervalMinutes` apart (i.e. genuinely
// back-to-back on the tee sheet, not just N open times somewhere in the day -
// these APIs only return slots with room, so a gap bigger than the interval
// means a fully-booked time sits between them).
function findConsecutiveWindow_(slots, intervalMinutes, playersNeeded, minCount, earliestMinutes) {
  var qualifying = slots
    .filter(function(s) {
      return s.availableSpots >= playersNeeded && timeStrToMinutes_(s.time) >= earliestMinutes;
    })
    .sort(function(a, b) { return timeStrToMinutes_(a.time) - timeStrToMinutes_(b.time); });

  var run = [];
  for (var i = 0; i < qualifying.length; i++) {
    if (run.length === 0) {
      run = [qualifying[i]];
    } else {
      var gap = timeStrToMinutes_(qualifying[i].time) - timeStrToMinutes_(run[run.length - 1].time);
      run = (gap === intervalMinutes) ? run.concat([qualifying[i]]) : [qualifying[i]];
    }
    if (run.length >= minCount) return run.slice(0, minCount);
  }
  return null;
}