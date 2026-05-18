/**
 * Updates the secondary disability column in all campus spreadsheets to allow
 * custom free-text entries in addition to the standard dropdown options.
 *
 * For monthly tabs (AUGUST–MARCH):
 *   - Sets the column G row 2 header to "List any Additional Disabilities"
 *   - Replaces the strict dropdown in G3:G with a "show warning" validation
 *     that still suggests options but permits custom values
 *
 * For the APRIL/MAY PROJECTIONS tab:
 *   - Sets the column H row 3 header to "List any Additional Disabilities"
 *   - Applies the same flexible validation to H4:H
 *
 * Reads campus spreadsheet IDs from CampusBMCSheetInfo column E.
 * Uses LockService to prevent concurrent runs.
 *
 * @returns {void}
 */
function modifyDataValidationToAllowCustomEntries() {
  var lock = LockService.getScriptLock();
  try {
    // Try to acquire the lock for 30 seconds
    lock.waitLock(30000);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Another instance is already running. Please wait and try again.');
    return;
  }
  
  try {
    var timestamp = new Date().toISOString();
    Logger.log('FUNCTION START: modifyDataValidationToAllowCustomEntries at ' + timestamp);
    
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
    
    // Get campus data (columns B=campus, E=spreadsheetId)
    var data = infoSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    // Define month sheets to update
    var monthlyTabs = ['AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER', 'JANUARY', 'FEBRUARY', 'MARCH'];
    var projectionTab = 'APRIL/ MAY PROJECTIONS';
    
    // Create data validation that allows custom entries with warnings
    var disabilityOptions = ['AU', 'OHI', 'ED', 'SI', 'ID', 'LD', 'OTHER'];
    var flexibleValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(disabilityOptions, true)
      .setAllowInvalid(true)  // This allows custom entries with a warning
      .setHelpText('Select one from the dropdown or if more, enter text listing the multiple disabilities. Valid options: ' + disabilityOptions.join(', '))
      .build();
    
    var updated = [];
    var skipped = [];
    var errors = [];
    
    // Process each campus
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
        Logger.log('Processing campus: ' + campus + ' (ID: ' + spreadsheetId + ')');
        var campusSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
        var campusUpdated = false;
        
        // Process monthly tabs (AUGUST through MARCH)
        for (var j = 0; j < monthlyTabs.length; j++) {
          var monthName = monthlyTabs[j];
          var sheet = campusSpreadsheet.getSheetByName(monthName);
          
          if (sheet) {
            try {
              // Update header in G2 to "List any Additional Disabilities"
              sheet.getRange('G2').setValue('List any Additional Disabilities');
              
              // Update data validation in G3:G to allow custom entries with warnings
              var lastRowInSheet = Math.max(sheet.getLastRow(), 1000);
              var validationRange = sheet.getRange('G3:G' + lastRowInSheet);
              validationRange.setDataValidation(flexibleValidation);
              
              Logger.log('  Updated header G2 and data validation G3:G in month tab: ' + monthName);
              campusUpdated = true;
            } catch (monthError) {
              errors.push('Row ' + (i + 2) + ' (' + campus + ') - ' + monthName + ': ' + monthError.toString());
            }
          } else {
            Logger.log('  Missing month tab: ' + monthName + ' in ' + campus);
          }
        }
        
        // Process APRIL/ MAY PROJECTIONS tab
        var projectionSheet = campusSpreadsheet.getSheetByName(projectionTab);
        if (projectionSheet) {
          try {
            // Update header in H3 to "List any Additional Disabilities"
            projectionSheet.getRange('H3').setValue('List any Additional Disabilities');
            
            // Update data validation in H4:H to allow custom entries with warnings
            var lastRowInProjection = Math.max(projectionSheet.getLastRow(), 1000);
            var projectionValidationRange = projectionSheet.getRange('H4:H' + lastRowInProjection);
            projectionValidationRange.setDataValidation(flexibleValidation);
            
            Logger.log('  Updated header H3 and data validation H4:H in projection tab');
            campusUpdated = true;
          } catch (projectionError) {
            errors.push('Row ' + (i + 2) + ' (' + campus + ') - ' + projectionTab + ': ' + projectionError.toString());
          }
        } else {
          Logger.log('  Missing projection tab: ' + projectionTab + ' in ' + campus);
        }
        
        if (campusUpdated) {
          updated.push(campus);
          Logger.log('Successfully processed campus: ' + campus);
        }
        
      } catch (campusError) {
        errors.push('Row ' + (i + 2) + ' (' + campus + '): ' + campusError.toString());
        Logger.log('Error processing campus ' + campus + ': ' + campusError.toString());
      }
    }
    
    // Show results to user
    var ui = SpreadsheetApp.getUi();
    var message = '';
    
    if (updated.length > 0) {
      message += 'Successfully updated headers and data validation to allow custom entries in ' + updated.length + ' campus(es):\n';
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
    ui.alert('Data Validation Update Complete', message, ui.ButtonSet.OK);
    
    Logger.log('FUNCTION END: modifyDataValidationToAllowCustomEntries at ' + new Date().toISOString());
    Logger.log('Summary - Updated: ' + updated.length + ', Skipped: ' + skipped.length + ', Errors: ' + errors.length);
    
  } catch (error) {
    Logger.log('CRITICAL ERROR in modifyDataValidationToAllowCustomEntries: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'A critical error occurred: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}
