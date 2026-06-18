function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Golf Outing')
    .addItem('— THURSDAY —', 'doNothing')
    .addItem('1. Send Thursday Invite', 'showThursdayForm')
    .addItem('2. Prepare Thursday Pairing Sheet', 'prepareThursdayPairingSheet')
    .addItem('3. Generate Thursday Draft Pairings', 'generateDraftPairings')
    .addItem('4. Send Thursday Pairings Email', 'sendPairingsEmail')
    .addSeparator()
    .addItem('— MONDAY —', 'doNothing')
    .addItem('1. Send Monday Invite', 'showMondayForm')
    .addItem('2. Prepare Monday Pairing Sheet', 'prepareMondayRSVPSheet')
    .addItem('3. Generate Monday Draft Pairings', 'generateMondayDraftPairings')
    .addItem('4. Send Monday Pairings Email', 'sendMondayPairingsEmail')
    .addSeparator()
    .addItem('Sync Player Names to Forms', 'syncFormDropdowns')
    .addItem('Sync Contact Groups', 'syncContactGroups')
    .addToUi();
}

function doNothing() {
  // Intentionally empty — used for section label menu items
}

function showThursdayForm() {
  var html = HtmlService.createHtmlOutputFromFile('ThursdayForm')
    .setWidth(450)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Thursday Golf Outing');
}

function showMondayForm() {
  var html = HtmlService.createHtmlOutputFromFile('MondayForm')
    .setWidth(450)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Monday Golf Outing');
}

function getCourseList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var coursesSheet = ss.getSheetByName("Courses");
  var courseData = coursesSheet.getRange(2, 1, coursesSheet.getLastRow() - 1, 1).getValues();
  return courseData.map(function(row) { return row[0]; }).filter(function(name) { return name !== ""; });
}