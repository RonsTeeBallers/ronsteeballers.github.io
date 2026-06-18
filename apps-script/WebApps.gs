function doGet(e) {
  var params = e.parameter;
  var action = params.action || '';
  var callback = params.callback || '';
  var result;

  if (action === 'getEvent') {
    result = getEventData(params.eventId, params.player);
  } else if (action === 'getConfirmed') {
    result = getConfirmedPlayers(params.eventId);
  } else if (action === 'getOpenEvents') {
    result = getOpenEvents();
  } else if (action === 'getVenues') {
    result = getVenues();
  } else if (action === 'createEvent') {
    result = createEvent(params);
  } else if (action === 'savePairings') {
    result = savePairings(params);
  } else if (action === 'getEventInfo') {
    result = getEventInfo(params.eventId);
  } else if (action === 'sendInviteEmails') {
    result = sendInviteEmails(params);
  } else if (action === 'submitRSVP') {
    result = submitRSVP(params);
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

  return ContentService.createTextOutput(JSON.stringify({error: 'Unknown action'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getEventData(eventId, playerSlug) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var eventsSheet = ss.getSheetByName('Events');
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

    // Find event by ID (date string)
    var eventsData = eventsSheet.getDataRange().getValues();
    var event = null;

    for (var i = 1; i < eventsData.length; i++) {
      var eventDate = eventsData[i][0];
      var dateStr = '';
      if (eventDate instanceof Date) {
        dateStr = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = eventDate.toString().trim();
      }
      if (dateStr === eventId && eventsData[i][5] === 'Open') {
        event = {
          date: dateStr,
          venueName: eventsData[i][1].toString(),
          time: eventsData[i][2] instanceof Date
            ? Utilities.formatDate(eventsData[i][2], Session.getScriptTimeZone(), 'h:mm a')
            : eventsData[i][2].toString(),
          slotsReserved: eventsData[i][3].toString(),
          mailingList: eventsData[i][4].toString(),
          notes: eventsData[i][6] ? eventsData[i][6].toString() : ''
        };
        break;
      }
    }

    if (!event) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found or not open'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
        title: 'Golf Outing — ' + event.venueName,
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
    var events = [];

    for (var i = 1; i < eventsData.length; i++) {
      var status = eventsData[i][5] ? eventsData[i][5].toString().trim() : '';
      if (status !== 'Open') continue;

      var eventDate = eventsData[i][0];
      var dateStr = eventDate instanceof Date
        ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : eventDate.toString().trim();

      var slotsReserved = parseInt(eventsData[i][3]) || 0;
      var totalSlots = slotsReserved * 4;

      // Count confirmed players for this specific event
      var formData = formSheet.getDataRange().getValues();
      var responseMap = {};
      for (var k = 1; k < formData.length; k++) {
        var name = formData[k][1].toString().trim();
        var response = formData[k][2].toString().trim();
        var rawEventId = formData[k][6];
      var rowEventId = '';
      if (rawEventId instanceof Date) {
        rowEventId = Utilities.formatDate(rawEventId, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (rawEventId) {
        rowEventId = rawEventId.toString().trim();
      }
        if (name === '') continue;
        if (rowEventId === '' || rowEventId === dateStr) {
          responseMap[name] = response;
        }
      }
      var confirmed = 0;
      for (var name in responseMap) {
        if (responseMap[name] === 'Yes') confirmed++;
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
    var playersSheet = ss.getSheetByName('Players');
    var eventsSheet = ss.getSheetByName('Events');

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

    var playersData = playersSheet.getDataRange().getValues();
    var preferredMap = {};
    var avoidMap = {};
    for (var i = 1; i < playersData.length; i++) {
      var name = playersData[i][0].toString().trim();
      preferredMap[name] = playersData[i][7] ? playersData[i][7].toString().trim() : '';
      avoidMap[name] = playersData[i][8] ? playersData[i][8].toString().trim() : '';
    }

    var formData = formSheet.getDataRange().getValues();
    var responseMap = {};

    for (var i = 1; i < formData.length; i++) {
      var name = formData[i][1].toString().trim();
      var response = formData[i][2].toString().trim();
      var walkRide = formData[i][3].toString().trim();
      var rawEventId = formData[i][6];
      var rowEventId = '';
      if (rawEventId instanceof Date) {
        rowEventId = Utilities.formatDate(rawEventId, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (rawEventId) {
        rowEventId = rawEventId.toString().trim().replace('EVT-', '');
      }
      if (name === '') continue;
      if (rowEventId === eventId) {
        responseMap[name] = {
          playing: response,
          walkRide: walkRide || 'No preference',
          comments: formData[i][5] ? formData[i][5].toString().trim() : '',
          scoring: formData[i][4] ? formData[i][4].toString().trim() : 'No Preference'
        };
      }
    }

    var confirmed = [];
    for (var name in responseMap) {
      if (responseMap[name].playing === 'Yes') {
        confirmed.push({
          name: name,
          walkRide: responseMap[name].walkRide,
          comments: responseMap[name].comments || '',
          scoring: responseMap[name].scoring || 'No Preference',
          preferred: preferredMap[name] || '',
          avoid: avoidMap[name] || ''
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
function debugFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName('Form Responses');
  var formData = formSheet.getDataRange().getValues();
  
  for (var i = 1; i < formData.length; i++) {
    var name = formData[i][1].toString().trim();
    var rowEventId = formData[i][6] ? formData[i][6].toString().trim() : 'BLANK';
    Logger.log('Row ' + i + ': ' + name + ' | EventID: [' + rowEventId + '] | Length: ' + rowEventId.length);
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
    var eventsData = eventsSheet.getDataRange().getValues();
    var eventRow = null;
    for (var i = 1; i < eventsData.length; i++) {
      var eventDate = eventsData[i][0];
      var dateStr = eventDate instanceof Date
        ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : eventDate.toString().trim();
      if (dateStr === eventId) {
        eventRow = eventsData[i];
        break;
      }
    }

    if (!eventRow) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
      headerCell.setValue('Foursome ' + (idx + 1) + ' — ' + teeTimeStr);
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

    // Build email body
    var emailLines = [];
    foursomes.forEach(function(foursome, idx) {
      var fsMinutes = startMinutes + idx * teeTimeInterval;
      var fh = Math.floor(fsMinutes / 60);
      var fm = fsMinutes % 60;
      var fap = fh >= 12 ? 'PM' : 'AM';
      fh = fh % 12 || 12;
      var teeTimeStr = fh + ':' + (fm < 10 ? '0' + fm : fm) + ' ' + fap;
      emailLines.push('<p><strong>Foursome ' + (idx + 1) + ' — ' + teeTimeStr + '</strong></p>');
      foursome.forEach(function(player) {
        emailLines.push('<p style="margin:0;padding-left:20px;">' + player.name + ' (' + player.walkRide + ')</p>');
      });
      emailLines.push('<br>');
    });

    var htmlBody =
      '<p>Gentlemen,</p>' +
      '<p>Here are the pairings for ' + formattedDate + ' at ' + venueName + ':</p>' +
      emailLines.join('') +
      '<p>Tee times reserved under Ron Blanton.</p>' +
      '<p>See you on the course!</p>';

    var subject = 'Golf Pairings — ' + formattedDate + ' at ' + venueName;

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

    // Save to Pairing History
    foursomes.forEach(function(foursome) {
      var historyRow = [formattedDate];
      foursome.forEach(function(player) { historyRow.push(player.name); });
      while (historyRow.length < 5) historyRow.push('');
      historyRow.push(venueName);
      historySheet.appendRow(historyRow);
    });

    // Update event status to Closed
    for (var i = 1; i < eventsData.length; i++) {
      var eventDate2 = eventsData[i][0];
      var dateStr2 = eventDate2 instanceof Date
        ? Utilities.formatDate(eventDate2, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : eventDate2.toString().trim();
      if (dateStr2 === eventId) {
        eventsSheet.getRange(i + 1, 6).setValue('Closed');
        break;
      }
    }

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
    'TEST — Golf Pairings email',
    '<p>This is a test of the direct pairings send via Brevo.</p>' +
    '<p><strong>Foursome 1 — 11:00 AM</strong></p>' +
    '<p style="margin:0;padding-left:20px;">Test Player (Walk)</p>'
  );
  Logger.log(JSON.stringify(result));
}

function getEventInfo(eventId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var eventsSheet = ss.getSheetByName('Events');
    var coursesSheet = ss.getSheetByName('Courses');
    var eventsData = eventsSheet.getDataRange().getValues();

    for (var i = 1; i < eventsData.length; i++) {
      var eventDate = eventsData[i][0];
      var dateStr = eventDate instanceof Date
        ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : eventDate.toString().trim();

      if (dateStr === eventId) {
        var venueName = eventsData[i][1].toString();
        var startTime = eventsData[i][2];
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
      }
    }

    return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function sendInviteEmails(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var playersSheet = ss.getSheetByName('Players');
    var eventsSheet = ss.getSheetByName('Events');
    var coursesSheet = ss.getSheetByName('Courses');

    var eventId = params.eventId;
    var mailingList = params.mailingList;
    var comment = params.comment || '';

    // Get event details
    var eventsData = eventsSheet.getDataRange().getValues();
    var eventRow = null;
    for (var i = 1; i < eventsData.length; i++) {
      var eventDate = eventsData[i][0];
      var dateStr = eventDate instanceof Date
        ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : eventDate.toString().trim();
      if (dateStr === eventId) {
        eventRow = eventsData[i];
        break;
      }
    }

    if (!eventRow) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Event not found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
    for (var c = 1; c < coursesData.length; c++) {
      if (coursesData[c][0].toString() === venueName) {
        var green = parseFloat(coursesData[c][7]) || 0;
        var cart = parseFloat(coursesData[c][8]) || 0;
        greenFee = green > 0 ? '$' + green.toFixed(2) : '';
        cartFee = cart > 0 ? '$' + cart.toFixed(2) : '';
        total = (green + cart) > 0 ? '$' + (green + cart).toFixed(2) : '';
        venueUrl = coursesData[c][6] ? coursesData[c][6].toString() : '';
        break;
      }
    }

    // Get players for this mailing list
    var playersData = playersSheet.getDataRange().getValues();
    var recipients = [];
    for (var p = 1; p < playersData.length; p++) {
      var active = playersData[p][4].toString().trim();
      var mondayPlayer = playersData[p][5].toString().trim();
      var specialInvitations = playersData[p][5].toString().trim();
      var email = playersData[p][3].toString().trim();
      var firstName = playersData[p][1].toString().trim();
      var lastName = playersData[p][2].toString().trim();
      var lastFirst = playersData[p][0].toString().trim();

      if (!email || email.indexOf('@') === -1) continue;

      var include = false;
      if (mailingList === 'Main Group' && active === 'Yes') include = true;
      if (mailingList === 'Monday' && (mondayPlayer === 'Yes')) include = true;
      if (mailingList === 'Indoor' && specialInvitations === 'Indoor') include = true;

      if (include) {
        var slug = lastFirst.toLowerCase()
          .replace(/,\s*/g, '-')
          .replace(/\s+/g, '-')
          .replace(/[^a-z0-9-]/g, '');

        recipients.push({
          firstName: firstName,
          email: email,
          slug: slug
        });
      }
    }

    if (recipients.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({error: 'No recipients found for mailing list: ' + mailingList}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Get Resend API key from Script Properties
    var apiKey = PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY');
    if (!apiKey) {
      return ContentService.createTextOutput(JSON.stringify({error: 'Brevo API key not configured (Script Property BREVO_API_KEY)'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var baseUrl = 'https://ronsteeballers.github.io';
    var signupUrl = baseUrl + '/signup.html?event=' + eventId;
    var sent = 0;
    var errors = [];

    // Send personalized email to each recipient
    recipients.forEach(function(recipient) {
      var rsvpUrl = baseUrl + '/rsvp.html?event=' + eventId + '&player=' + recipient.slug;

      var feeSection = '';
      if (total) {
        feeSection = '<p style="color:#5d6d7e;font-size:14px;">Green Fee: ' + greenFee +
          ' &nbsp;|&nbsp; Cart: ' + cartFee +
          ' &nbsp;|&nbsp; <strong>Total: ' + total + '</strong></p>';
      }

      var venueLink = venueUrl
        ? '<a href="' + venueUrl + '" style="color:#1a5276;">' + venueName + '</a>'
        : venueName;

      var commentSection = comment
        ? '<p style="background:#eaf2ff;padding:12px;border-radius:8px;border-left:3px solid #1a5276;">' + comment + '</p>'
        : '';

      var htmlBody =
        '<div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;max-width:600px;margin:0 auto;padding:20px;">' +
        '<div style="background:#1a5276;color:white;padding:20px;border-radius:12px 12px 0 0;text-align:center;">' +
        '<div style="font-size:32px;">⛳</div>' +
        '<h1 style="margin:8px 0;font-size:22px;">Golf Outing Invitation</h1>' +
        '</div>' +
        '<div style="background:white;padding:24px;border:1px solid #e0e8f0;border-top:none;">' +
        '<p style="font-size:18px;font-weight:700;color:#1a2332;">Hi ' + recipient.firstName + '!</p>' +
        '<p style="color:#5d6d7e;">You\'re invited to join us for golf:</p>' +
        '<div style="background:#f0f4f8;padding:16px;border-radius:8px;margin:16px 0;">' +
        '<p style="margin:4px 0;font-size:16px;font-weight:700;color:#1a2332;">📅 ' + formattedDate + '</p>' +
        '<p style="margin:4px 0;font-size:16px;color:#1a2332;">📍 ' + venueLink + '</p>' +
        '<p style="margin:4px 0;font-size:16px;color:#1a2332;">⏰ ' + timeStr + '</p>' +
        '<p style="margin:4px 0;font-size:14px;color:#5d6d7e;">' + slotsReserved + ' tee times reserved</p>' +
        '</div>' +
        feeSection +
        commentSection +
        '<div style="text-align:center;margin:24px 0;">' +
        '<a href="' + rsvpUrl + '" style="display:inline-block;background:#1a5276;color:white;padding:16px 32px;border-radius:10px;text-decoration:none;font-size:18px;font-weight:700;margin:0 8px;">✅ I\'M IN</a>' +
        '<a href="' + rsvpUrl + '" style="display:inline-block;background:#922b21;color:white;padding:16px 32px;border-radius:10px;text-decoration:none;font-size:18px;font-weight:700;margin:0 8px;">❌ I\'M OUT</a>' +
        '</div>' +
        '<p style="text-align:center;color:#5d6d7e;font-size:14px;">See who else is playing: <a href="' + signupUrl + '" style="color:#1a5276;">View Signup List</a></p>' +
        '</div>' +
        '<div style="background:#f0f4f8;padding:12px;border-radius:0 0 12px 12px;text-align:center;">' +
        '<p style="color:#aab7c4;font-size:12px;margin:0;">See you on the course!</p>' +
        '</div>' +
        '</div>';

      try {
        var response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
          method: 'post',
          headers: {
            'api-key': apiKey,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify({
            sender: { name: 'RonsTeeBallers', email: 'ronsteeballers@gmail.com' },
            to: [{ email: recipient.email, name: recipient.firstName }],
            subject: 'Golf Outing — ' + formattedDate + ' at ' + venueName,
            htmlContent: htmlBody
          }),
          muteHttpExceptions: true
        });

        var responseCode = response.getResponseCode();
        if (responseCode === 200 || responseCode === 201) {
          sent++;
        } else {
          errors.push(recipient.email + ': ' + response.getContentText());
        }
      } catch(err) {
        errors.push(recipient.email + ': ' + err.message);
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
function submitRSVP(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName('Form Responses');

    var timestamp = new Date();
    var row = [
      timestamp,
      params.playerName || '',
      params.playing || '',
      params.walkRide || '',
      params.scoring || '',
      params.comments || '',
      params.eventId || ''
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
function testTextStorage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form Responses');
  var lastRow = sheet.getLastRow() + 1;
  
  // Method 1 - apostrophe prefix
  sheet.getRange(lastRow, 6).setValue("'" + '2026-06-19');
  
  Logger.log('Cell type: ' + typeof sheet.getRange(lastRow, 6).getValue());
  Logger.log('Cell value: ' + sheet.getRange(lastRow, 6).getValue());
}