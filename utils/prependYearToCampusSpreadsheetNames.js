/**
 * Prepends '25-26 ' to the name of each campus spreadsheet listed in CampusBMCSheetInfo (column E).
 * Skips if the spreadsheet name already starts with '25-26 '.
 * Alerts the user with a summary of changes and errors.
 */
function prependYearToCampusSpreadsheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var infoSheet = ss.getSheetByName('CampusBMCSheetInfo');
  if (!infoSheet) {
    SpreadsheetApp.getUi().alert('CampusBMCSheetInfo sheet not found.');
    return;
  }
  var lastRow = infoSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows in CampusBMCSheetInfo.');
    return;
  }
  var data = infoSheet.getRange(2, 5, lastRow - 1, 1).getValues(); // E2:E
  var changed = [];
  var skipped = [];
  var errors = [];
  for (var i = 0; i < data.length; i++) {
    var spreadsheetId = (data[i][0] || '').toString().trim();
    if (!spreadsheetId) continue;
    try {
      var file = DriveApp.getFileById(spreadsheetId);
      var name = file.getName();
      if (name.indexOf('25-26 ') === 0) {
        skipped.push(name);
        continue;
      }
      var newName = '25-26 ' + name;
      file.setName(newName);
      changed.push(newName);
    } catch (e) {
      errors.push('Row ' + (i+2) + ': ' + spreadsheetId + ' - ' + e);
    }
  }
  var msg = 'Campus spreadsheet renaming complete.\n';
  if (changed.length) msg += '\nRenamed (' + changed.length + '):\n' + changed.join('\n');
  if (skipped.length) msg += '\n\nAlready correct (' + skipped.length + '):\n' + skipped.join('\n');
  if (errors.length) msg += '\n\nErrors (' + errors.length + '):\n' + errors.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}
