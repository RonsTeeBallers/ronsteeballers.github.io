function syncFormDropdowns() {
  // Get active player names from Players tab
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var playersSheet = ss.getSheetByName("Players");
  var playersData = playersSheet.getDataRange().getValues();
  
  // Get all active players (Full Name, Active = Yes)
  var activeNames = playersData
    .slice(1)
    .filter(function(row) { return row[4] === "Yes"; })
    .map(function(row) { return row[0]; })
    .filter(function(name) { return name !== ""; })
    .sort();

  // Get Monday players (Full Name, Monday Player = Yes)
  var mondayNames = playersData
    .slice(1)
    .filter(function(row) { return row[5] === "Yes"; })
    .map(function(row) { return row[0]; })
    .filter(function(name) { return name !== ""; })
    .sort();

  
  
  // Update Thursday form
  updateFormDropdown(
    "1kTbC5H-S7UjcF1A3ZRHTx1F8Gsq7xtopIer-u0lXyq4",
    activeNames
  );

  // Update Monday form
  updateFormDropdown(
    "14Yts9Ipo-bI2jTmJcVg5EmB8uoaEIW399e7naBvvzEk",
    mondayNames
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Form dropdowns updated successfully!",
    "Sync Complete",
    5
  );
}

function updateFormDropdown(formId, names) {
  var form = FormApp.openById(formId);
  var items = form.getItems();
  
  // Find the Your Name question
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === "Your Name") {
      var listItem = items[i].asListItem();
      var choices = names.map(function(name) {
        return listItem.createChoice(name);
      });
      listItem.setChoices(choices);
      break;
    }
  }
}