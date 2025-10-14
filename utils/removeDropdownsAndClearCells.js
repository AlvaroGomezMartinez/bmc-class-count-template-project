/**
 * Removes dropdown data validation and clears cells for specified ranges in each campus spreadsheet.
 *
 * For monthly tabs (AUGUST - MARCH): removes data validation from G3:G (does NOT clear cell contents).
 * For APRIL/ MAY PROJECTIONS: removes data validation from H4:H (does NOT clear cell contents).
 *
 * Reads campus spreadsheet IDs from `CampusBMCSheetInfo` column E (same data block used by the other function).
 */
function removeDropdowns() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Another instance is already running. Please wait and try again.');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var infoSheet = ss.getSheetByName('CampusBMCSheetInfo');
    if (!infoSheet) {
      throw new Error('CampusBMCSheetInfo sheet not found');
    }

    var lastRow = infoSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No data rows in CampusBMCSheetInfo.');
      return;
    }

    var data = infoSheet.getRange(2, 1, lastRow - 1, 5).getValues();

    var monthlyTabs = ['AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JANUARY', 'FEBRUARY', 'MARCH'];
    var projectionTab = 'APRIL/ MAY PROJECTIONS';

    var updated = [];
    var skipped = [];
    var errors = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var campus = row[1]; // Column B
      var spreadsheetId = row[4]; // Column E

      if (!campus) {
        skipped.push('Row ' + (i + 2) + ': Missing campus name');
        continue;
      }

      if (!spreadsheetId) {
        skipped.push('Row ' + (i + 2) + ' (' + campus + '): Missing spreadsheet ID');
        continue;
      }

      try {
        Logger.log('Processing (remove) campus: ' + campus + ' (ID: ' + spreadsheetId + ')');
        var campusSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
        var campusUpdated = false;

        // Monthly tabs: remove validation and clear G3:G
        for (var j = 0; j < monthlyTabs.length; j++) {
          var monthName = monthlyTabs[j];
          var sheet = campusSpreadsheet.getSheetByName(monthName);
          if (sheet) {
            try {
              var lastRowInSheet = Math.max(sheet.getLastRow(), 1000);
              var rangeA = sheet.getRange('G3:G' + lastRowInSheet);
              rangeA.clearDataValidations();
              Logger.log('  Cleared G3:G in ' + monthName + ' for ' + campus);
              campusUpdated = true;
            } catch (monthError) {
              errors.push('Row ' + (i + 2) + ' (' + campus + ') - ' + monthName + ': ' + monthError.toString());
            }
          } else {
            Logger.log('  Missing month tab: ' + monthName + ' in ' + campus);
          }
        }

        // Projection tab: remove validation and clear H4:H
        var projectionSheet = campusSpreadsheet.getSheetByName(projectionTab);
        if (projectionSheet) {
            try {
            var lastRowInProjection = Math.max(projectionSheet.getLastRow(), 1000);
            var rangeB = projectionSheet.getRange('H4:H' + lastRowInProjection);
            rangeB.clearDataValidations();
            Logger.log('  Cleared H4:H in ' + projectionTab + ' for ' + campus);
            campusUpdated = true;
          } catch (projectionError) {
            errors.push('Row ' + (i + 2) + ' (' + campus + ') - ' + projectionTab + ': ' + projectionError.toString());
          }
        } else {
          Logger.log('  Missing projection tab: ' + projectionTab + ' in ' + campus);
        }

        if (campusUpdated) {
          updated.push(campus);
        }

      } catch (campusError) {
        errors.push('Row ' + (i + 2) + ' (' + campus + '): ' + campusError.toString());
      }
    }

    var ui = SpreadsheetApp.getUi();
    var message = '';

    if (updated.length > 0) {
      message += 'Successfully removed dropdowns and cleared cells in ' + updated.length + ' campus(es):\n';
      message += '• ' + updated.join('\n• ') + '\n\n';
    }

    if (skipped.length > 0) {
      message += 'Skipped ' + skipped.length + ' row(s):\n';
      message += '• ' + skipped.join('\n• ') + '\n\n';
    }

    if (errors.length > 0) {
      message += 'Encountered ' + errors.length + ' error(s):\n';
      message += '• ' + errors.join('\n• ') + '\n\n';
    }

    if (message === '') {
      message = 'No updates were made. Check that there are valid campus spreadsheet IDs in column E of CampusBMCSheetInfo.';
    }

    message += '\nCheck the execution logs for detailed information.';
    ui.alert('Remove Dropdowns Complete', message, ui.ButtonSet.OK);

    Logger.log('FUNCTION END: removeDropdownsAndClearCells at ' + new Date().toISOString());
    Logger.log('Summary - Updated: ' + updated.length + ', Skipped: ' + skipped.length + ', Errors: ' + errors.length);

  } catch (error) {
    Logger.log('CRITICAL ERROR in removeDropdownsAndClearCells: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'A critical error occurred: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}
