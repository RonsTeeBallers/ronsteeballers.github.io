function generateDraftPairings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rsvpSheet = ss.getSheetByName("Thursday RSVPs");
  var draftSheet = ss.getSheetByName("Draft Pairings");
  var eventsSheet = ss.getSheetByName("Events");
  var coursesSheet = ss.getSheetByName("Courses");

  // Clear existing draft
  draftSheet.clearContents();

  // Get Thursday RSVPs data
  var rsvpData = rsvpSheet.getDataRange().getValues();

  // Build foursome groups from column A assignments
  var foursomeMap = {};
  var unassigned = [];

  for (var i = 1; i < rsvpData.length; i++) {
    var foursomeNum = rsvpData[i][0].toString().trim();
    var name = rsvpData[i][1].toString().trim();
    var walkRide = rsvpData[i][2].toString().trim() || "No preference";

    if (name === "") continue;

    if (foursomeNum === "" || foursomeNum === "0") {
      unassigned.push({
        displayName: name,
        walkRide: walkRide
      });
    } else {
      if (!foursomeMap[foursomeNum]) {
        foursomeMap[foursomeNum] = [];
      }
      foursomeMap[foursomeNum].push({
        displayName: name,
        walkRide: walkRide
      });
    }
  }

  // Sort foursome numbers numerically
  var foursomeNumbers = Object.keys(foursomeMap).sort(function(a, b) {
    return parseFloat(a) - parseFloat(b);
  });

  // Get event details for tee times
  var eventData = eventsSheet.getDataRange().getValues();
  var startingTeeTime = "";
  var eventCourseName = "";

  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Thursday") {
      startingTeeTime = eventData[i][2];
      eventCourseName = eventData[i][1];
      break;
    }
  }

  // Get tee time interval from Courses tab
  var courseData = coursesSheet.getDataRange().getValues();
  var teeTimeInterval = 9;

  for (var c = 1; c < courseData.length; c++) {
    if (courseData[c][0] === eventCourseName) {
      var interval = parseInt(courseData[c][10]);
      if (!isNaN(interval) && interval > 0) {
        teeTimeInterval = interval;
      }
      break;
    }
  }

  // Parse starting tee time
  var startMinutes = 0;
  if (startingTeeTime instanceof Date) {
    startMinutes = startingTeeTime.getHours() * 60 + startingTeeTime.getMinutes();
  } else if (typeof startingTeeTime === "string" && startingTeeTime.indexOf(":") !== -1) {
    var timeParts = startingTeeTime.split(":");
    startMinutes = parseInt(timeParts[0]) * 60 + parseInt(timeParts[1]);
  }

  // Write foursomes to Draft Pairings tab
  var row = 1;

  foursomeNumbers.forEach(function(num, index) {
    var foursome = foursomeMap[num];

    // Calculate tee time
    var foursomeMinutes = startMinutes + (index * teeTimeInterval);
    var hours = Math.floor(foursomeMinutes / 60);
    var minutes = foursomeMinutes % 60;
    var ampm = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    var teeTimeStr = hours + ":" + (minutes < 10 ? "0" + minutes : minutes) + " " + ampm;

    // Write header
    var headerCell = draftSheet.getRange(row, 1);
    headerCell.setValue("Foursome " + num + " — " + teeTimeStr);
    headerCell.setFontWeight("bold");
    row++;

    // Write players
    foursome.forEach(function(player) {
      var playerCell = draftSheet.getRange(row, 1);
      playerCell.setValue("   " + player.displayName + " (" + player.walkRide + ")");
      playerCell.setFontWeight("normal");
      row++;
    });

    // Blank row between foursomes
    draftSheet.getRange(row, 1).setFontWeight("normal");
    row++;
  });

  // Write unassigned players at the bottom if any
  if (unassigned.length > 0) {
    var unassignedHeader = draftSheet.getRange(row, 1);
    unassignedHeader.setValue("Unassigned Players");
    unassignedHeader.setFontWeight("bold");
    row++;

    unassigned.forEach(function(player) {
      var playerCell = draftSheet.getRange(row, 1);
      playerCell.setValue("   " + player.displayName + " (" + player.walkRide + ")");
      playerCell.setFontWeight("normal");
      row++;
    });
  }

  ss.toast("Draft pairings generated from RSVP tab assignments!", "Pairings", 5);
}

function getPlayerDetails(lastFirst, playersSheet) {
  var data = playersSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === lastFirst) {
      return {
        displayName: data[i][1] + " " + data[i][2],
        pairingNotes: "",
        alwaysTogether: data[i][6] ? data[i][6].toString().split(";").map(function(s){return s.trim().toLowerCase();}) : [],
        preferred: data[i][7] ? data[i][7].toString().split(";").map(function(s){return s.trim().toLowerCase();}) : [],
        avoid: data[i][8] ? data[i][8].toString().split(";").map(function(s){return s.trim().toLowerCase();}) : []
      };
    }
  }
  var parts = lastFirst.split(",");
  return {
    displayName: parts.length > 1 ? parts[1].trim() + " " + parts[0].trim() : lastFirst,
    pairingNotes: "",
    alwaysTogether: [],
    preferred: [],
    avoid: []
  };
}
function buildFoursomes(players) {
  var foursomes = [];
  var assigned = [];

// Step 1 — Handle Always Together groups first
  // Only lock players together if their partner is actually playing this week
players.forEach(function(player) {
    if (assigned.map(function(n){return n.toLowerCase();}).indexOf(player.name.toLowerCase()) !== -1) return;
    if (player.alwaysTogether.length > 0) {
      var group = [player];
      player.alwaysTogether.forEach(function(partnerName) {
        var partner = players.find(function(p) { return p.name.toLowerCase() === partnerName; });
        if (partner && assigned.indexOf(partner.name) === -1 && group.length < 4) {
          group.push(partner);
        }
      });
      // Only lock the group if at least one partner was found
      if (group.length > 1) {
        group.forEach(function(p) { assigned.push(p.name); });
        foursomes.push(group);
      }
    }
  });

  // Step 2 — Handle guests — keep with their host
players.forEach(function(player) {
    if (assigned.map(function(n){return n.toLowerCase();}).indexOf(player.name.toLowerCase()) !== -1) return;
    if (!player.guest) return;
    var hostFoursome = null;
    foursomes.forEach(function(foursome) {
      foursome.forEach(function(member) {
        if (player.guest.toLowerCase().indexOf(member.displayName.toLowerCase()) !== -1) {
          hostFoursome = foursome;
        }
      });
    });
    if (hostFoursome && hostFoursome.length < 4) {
      hostFoursome.push(player);
      assigned.push(player.name);
    }
  });

// Step 3 — Get all remaining unassigned players (case insensitive check)
  var assignedLower = assigned.map(function(n) { return n.toLowerCase(); });
  var remaining = players.filter(function(p) {
    return assignedLower.indexOf(p.name.toLowerCase()) === -1;
  });

  // Step 4 — Apply preferred partner grouping on all remaining players together
  var preferredAssigned = [];
  var groups = [];

  remaining.forEach(function(player) {
    if (preferredAssigned.indexOf(player.name) !== -1) return;
    var group = [player];
    preferredAssigned.push(player.name);

    if (player.preferred.length > 0) {
      player.preferred.forEach(function(partnerName) {
        if (group.length >= 4) return;
        if (preferredAssigned.indexOf(partnerName) !== -1) return;
        var partner = remaining.find(function(p) {
          return p.name.toLowerCase() === partnerName;
        });
        if (partner) {
          var avoided = false;
          group.forEach(function(member) {
            if (member.avoid.indexOf(partnerName) !== -1) avoided = true;
            if (partner.avoid.indexOf(member.name.toLowerCase()) !== -1) avoided = true;
          });
          if (!avoided) {
            group.push(partner);
            preferredAssigned.push(partner.name);
          }
        }
      });
    }
    groups.push(group);
  });

// Step 5 — Flatten all groups into one pool — deduplicate by name
  var pool = [];
  var poolNames = [];
  groups.forEach(function(group) {
    group.forEach(function(player) {
      if (poolNames.indexOf(player.name.toLowerCase()) === -1) {
        pool.push(player);
        poolNames.push(player.name.toLowerCase());
      }
    });
  });

  // Step 6 — Fill foursomes from pool respecting avoid rules
  while (pool.length > 0) {
    var foursome = [pool[0]];
    var remaining_pool = [];

    for (var i = 1; i < pool.length; i++) {
      if (foursome.length >= 4) {
        remaining_pool.push(pool[i]);
        continue;
      }
      var player = pool[i];
      var avoided = false;
      foursome.forEach(function(member) {
        if (member.avoid.indexOf(player.name.toLowerCase()) !== -1) avoided = true;
        if (player.avoid.indexOf(member.name.toLowerCase()) !== -1) avoided = true;
      });
      if (!avoided) {
        foursome.push(player);
      } else {
        remaining_pool.push(player);
      }
    }

    foursomes.push(foursome);
    pool = remaining_pool;
  }

  return foursomes;
}

function buildPreferredGroups(players, size) {
  var groups = [];
  var assigned = [];

  players.forEach(function(player) {
    if (assigned.indexOf(player.name) !== -1) return;
    var group = [player];
    assigned.push(player.name);
    if (player.preferred.length > 0) {
      player.preferred.forEach(function(partnerName) {
        if (assigned.indexOf(partnerName) !== -1) return;
        if (group.length >= size) return;
        var partner = players.find(function(p) { return p.name.toLowerCase() === partnerName; });
        if (partner) {
          var avoided = false;
          group.forEach(function(member) {
            if (member.avoid.indexOf(partnerName) !== -1) avoided = true;
            if (partner.avoid.indexOf(member.name) !== -1) avoided = true;
          });
          if (!avoided) {
            group.push(partner);
            assigned.push(partner.name);
          }
        }
      });
    }
    groups.push(group);
  });

  var unassigned = players.filter(function(p) {
    return assigned.indexOf(p.name) === -1;
  });

  unassigned.forEach(function(player) {
    var added = false;
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].length < size) {
        var avoided = false;
        groups[i].forEach(function(member) {
          if (member.avoid.indexOf(player.name) !== -1) avoided = true;
          if (player.avoid.indexOf(member.name) !== -1) avoided = true;
        });
        if (!avoided) {
          groups[i].push(player);
          assigned.push(player.name);
          added = true;
          break;
        }
      }
    }
    if (!added) {
      groups.push([player]);
      assigned.push(player.name);
    }
  });

  var smallGroups = groups.filter(function(g) { return g.length < size; });
  var fullGroups = groups.filter(function(g) { return g.length >= size; });
  var consolidated = [];
  var currentGroup = [];

  smallGroups.forEach(function(group) {
    group.forEach(function(player) {
      currentGroup.push(player);
      if (currentGroup.length === size) {
        consolidated.push(currentGroup);
        currentGroup = [];
      }
    });
  });

  if (currentGroup.length > 0) {
    consolidated.push(currentGroup);
  }

  return fullGroups.concat(consolidated);
}
function writeDraftPairings(foursomes, draftSheet, startingTeeTime, teeTimeInterval) {
  var row = 1;

  var startMinutes = 0;
  if (startingTeeTime instanceof Date) {
    startMinutes = startingTeeTime.getHours() * 60 + startingTeeTime.getMinutes();
  } else if (typeof startingTeeTime === "string" && startingTeeTime.indexOf(":") !== -1) {
    var timeParts = startingTeeTime.split(":");
    startMinutes = parseInt(timeParts[0]) * 60 + parseInt(timeParts[1]);
  }

  foursomes.forEach(function(foursome, index) {
    var foursomeMinutes = startMinutes + (index * teeTimeInterval);
    var hours = Math.floor(foursomeMinutes / 60);
    var minutes = foursomeMinutes % 60;
    var ampm = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    var teeTimeStr = hours + ":" + (minutes < 10 ? "0" + minutes : minutes) + " " + ampm;

    var headerCell = draftSheet.getRange(row, 1);
    headerCell.setValue("Foursome " + (index + 1) + " — " + teeTimeStr);
    headerCell.setFontWeight("bold");
    row++;

    foursome.forEach(function(player) {
      var playerLine = "   " + player.displayName + " (" + player.walkRide + ")";
      var playerCell = draftSheet.getRange(row, 1);
      playerCell.setValue(playerLine);
      playerCell.setFontWeight("normal");
      row++;
    });

    draftSheet.getRange(row, 1).setFontWeight("normal");
    row++;
  });
}

function sendPairingsEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var draftSheet = ss.getSheetByName("Draft Pairings");
  var eventsSheet = ss.getSheetByName("Events");
  var historySheet = ss.getSheetByName("Pairing History");

  var eventData = eventsSheet.getDataRange().getValues();
  var eventRow = null;
  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Thursday") {
      eventRow = eventData[i];
      break;
    }
  }

  if (!eventRow) {
    ss.toast("No Thursday event found.", "Error", 5);
    return;
  }

  var eventDate = eventRow[0];
  var courseName = eventRow[1];
  var formattedDate = (eventDate instanceof Date)
    ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy")
    : eventDate;

  var draftData = draftSheet.getDataRange().getValues();
  var emailBody = "";
  var historyRows = [];
  var currentFoursome = [];

  draftData.forEach(function(row) {
    var line = row[0].toString();
    if (line.trim() === "") {
      if (currentFoursome.length > 0) {
        while (currentFoursome.length < 4) currentFoursome.push("");
        historyRows.push([formattedDate].concat(currentFoursome).concat([courseName]));
        currentFoursome = [];
      }
      emailBody += "\n";
    } else if (line.indexOf("Foursome") === 0) {
      emailBody += line + "\n";
    } else {
      var cleanLine = line.split(" — ")[0];
      emailBody += cleanLine + "\n";
      var playerName = cleanLine.replace(/\s+\(.*\)/, "").trim();
      currentFoursome.push(playerName);
    }
  });

  if (currentFoursome.length > 0) {
    while (currentFoursome.length < 4) currentFoursome.push("");
    historyRows.push([formattedDate].concat(currentFoursome).concat([courseName]));
  }

  var subject = "Golf Pairings — " + formattedDate + " at " + courseName;
  var emailBodyHtml = emailBody
    .split("\n")
    .map(function(line) {
      if (line.indexOf("Foursome") === 0) {
        return "<p><strong>" + line + "</strong></p>";
      } else if (line.trim() === "") {
        return "<br>";
      } else {
        return "<p style='margin:0; padding-left:20px;'>" + line.trim() + "</p>";
      }
    })
    .join("");

  var htmlBody =
    "<p>Gentlemen,</p>" +
    "<p>Here are the pairings for this week:</p>" +
    emailBodyHtml +
    "<p>See you on the course!<br>Ron</p>";

  GmailApp.createDraft("", subject, htmlBody, {htmlBody: htmlBody});

  historyRows.forEach(function(historyRow) {
    historySheet.appendRow(historyRow);
  });

  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Thursday") {
      eventsSheet.getRange(i + 1, 6).setValue("Closed");
      break;
    }
  }

  ss.toast("Pairings email drafted and saved to history!", "Pairings", 5);
}
function generateMondayDraftPairings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rsvpSheet = ss.getSheetByName("Monday RSVPs");
  var playersSheet = ss.getSheetByName("Players");
  var draftSheet = ss.getSheetByName("Draft Pairings");

  draftSheet.clearContents();

  var rsvpData = rsvpSheet.getDataRange().getValues();
  var players = [];

// Build players list deduplicating by name — last response wins
  var playerMap = {};
  for (var i = 1; i < rsvpData.length; i++) {
    var name = rsvpData[i][0];       // Column A — Player Name
    var walkRide = rsvpData[i][1] || "No preference";  // Column B — Walk/Ride
    var guestInfo = rsvpData[i][2];  // Column C — Comments
    if (name === "") continue;
    // Overwrite any previous entry — last response wins
    playerMap[name] = {
      walkRide: walkRide || "No preference",
      guestInfo: guestInfo || ""
    };
  }

  // Convert map to players array
  for (var name in playerMap) {
    var playerDetails = getPlayerDetails(name, playersSheet);
    players.push({
      name: name,
      displayName: playerDetails.displayName,
      walkRide: playerMap[name].walkRide,
      alwaysTogether: playerDetails.alwaysTogether,
      preferred: playerDetails.preferred,
      avoid: playerDetails.avoid,
      pairingNotes: playerDetails.pairingNotes,
      guest: playerMap[name].guestInfo
    });
  }

  if (players.length === 0) {
    ss.toast("No confirmed players found in Monday RSVPs tab.", "Error", 5);
    return;
  }

  var eventsSheet = ss.getSheetByName("Events");
  var eventData = eventsSheet.getDataRange().getValues();
  var startingTeeTime = "";
  var eventCourseName = "Indian Lakes";

  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Monday") {
      startingTeeTime = eventData[i][2];
      break;
    }
  }

  var coursesSheet = ss.getSheetByName("Courses");
  var courseData = coursesSheet.getDataRange().getValues();
  var teeTimeInterval = 9;

  for (var c = 1; c < courseData.length; c++) {
    if (courseData[c][0] === eventCourseName) {
      var interval = parseInt(courseData[c][10]);
      if (!isNaN(interval) && interval > 0) {
        teeTimeInterval = interval;
      }
      break;
    }
  }

  var foursomes = buildFoursomes(players);
  writeDraftPairings(foursomes, draftSheet, startingTeeTime, teeTimeInterval);
  ss.toast("Monday draft pairings generated!", "Pairings", 5);
}

function sendMondayPairingsEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var draftSheet = ss.getSheetByName("Draft Pairings");
  var eventsSheet = ss.getSheetByName("Events");
  var historySheet = ss.getSheetByName("Pairing History");
  var rsvpSheet = ss.getSheetByName("Monday RSVPs");

  var eventData = eventsSheet.getDataRange().getValues();
  var eventRow = null;
  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Monday") {
      eventRow = eventData[i];
      break;
    }
  }

  if (!eventRow) {
    ss.toast("No Monday event found.", "Error", 5);
    return;
  }

  var eventDate = eventRow[0];
  var courseName = "Indian Lakes Golf Course";
  var formattedDate = (eventDate instanceof Date)
    ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy")
    : eventDate;

  var draftData = draftSheet.getDataRange().getValues();
  var emailBody = "";
  var historyRows = [];
  var currentFoursome = [];

  draftData.forEach(function(row) {
    var line = row[0].toString();
    if (line.trim() === "") {
      if (currentFoursome.length > 0) {
        while (currentFoursome.length < 4) currentFoursome.push("");
        historyRows.push([formattedDate].concat(currentFoursome).concat([courseName]));
        currentFoursome = [];
      }
      emailBody += "\n";
    } else if (line.indexOf("Foursome") === 0) {
      emailBody += line + "\n";
    } else {
      var cleanLine = line.split(" — ")[0];
      emailBody += cleanLine + "\n";
      var playerName = cleanLine.replace(/\s+\(.*\)/, "").trim();
      currentFoursome.push(playerName);
    }
  });

  if (currentFoursome.length > 0) {
    while (currentFoursome.length < 4) currentFoursome.push("");
    historyRows.push([formattedDate].concat(currentFoursome).concat([courseName]));
  }

  var confirmedEmails = [];

  var emailBodyHtml = emailBody
    .split("\n")
    .map(function(line) {
      if (line.indexOf("Foursome") === 0) {
        return "<p><strong>" + line + "</strong></p>";
      } else if (line.trim() === "") {
        return "<br>";
      } else {
        return "<p style='margin:0; padding-left:20px;'>" + line.trim() + "</p>";
      }
    })
    .join("");

  var subject = "Golf Pairings — " + formattedDate + " at Indian Lakes";
  var htmlBody =
    "<p>Gentlemen,</p>" +
    "<p>Here are the pairings for this week:</p>" +
    emailBodyHtml +
    "<p>See you on the course!<br>Ron</p>";

  GmailApp.createDraft("", subject, htmlBody, {htmlBody: htmlBody});

  historyRows.forEach(function(historyRow) {
    historySheet.appendRow(historyRow);
  });

  for (var i = eventData.length - 1; i >= 1; i--) {
    if (eventData[i][4] === "Monday") {
      eventsSheet.getRange(i + 1, 6).setValue("Closed");
      break;
    }
  }

  ss.toast("Monday pairings email drafted and saved to history!", "Pairings", 5);
}