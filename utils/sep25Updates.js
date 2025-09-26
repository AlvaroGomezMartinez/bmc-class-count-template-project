/** This file contains updates that were requested from the user.
 * - Update the header on D2 to read, "2025/2026 Campus"
 * - Update the data verification lists in cols F & G to have the correct disability labels
 * - Update the header on P2 to have, "Home Campus"
 * - Add borders down column P
 */

/**
 * Updates each campus file with September 25, 2025 requirements.
 * 
 * Goes into each campus file listed in CampusBMCSheetInfo column E and updates:
 * - Monthly tabs (AUGUST-MARCH): D2="2025/2026 Campus", P2="Home Campus", borders in column P, 
 *   data validation F3:G with "AU, OHI, ED, SI, ID, LD, OTHER"
 * - APRIL/ MAY PROJECTIONS tab: data validation G4:H with "AU, OHI, ED, SI, ID, LD, OTHER"
 * 
 * @returns {void}
 */
function updateCampusFilesWithSep25Updates() {
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
    Logger.log('FUNCTION START: updateCampusFilesWithSep25Updates at ' + timestamp);
    
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
    
    // Create data validation with new disability labels
    var disabilityOptions = ['AU', 'OHI', 'ED', 'SI', 'ID', 'LD', 'OTHER'];
    var validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(disabilityOptions, true)
      .setAllowInvalid(false)
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
              // Update D2 with "2025/2026 Campus"
              sheet.getRange('D2').setValue('2025/2026 Campus');
              
              // Update P2 with "Home Campus"
              sheet.getRange('P2').setValue('Home Campus');
              
              // Add borders to column P (assuming rows 1-1000 for comprehensive coverage)
              var columnPRange = sheet.getRange('P:P');
              columnPRange.setBorder(true, true, true, true, true, true);
              
              // Update data validation for F3:G with new disability options
              // Determine the range - go to last row or reasonable default
              var lastRowInSheet = Math.max(sheet.getLastRow(), 1000);
              var validationRange = sheet.getRange('F3:G' + lastRowInSheet);
              validationRange.setDataValidation(validation);
              
              Logger.log('  Updated month tab: ' + monthName);
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
            // Update data validation for G4:H with new disability options
            var lastRowInProjection = Math.max(projectionSheet.getLastRow(), 1000);
            var projectionValidationRange = projectionSheet.getRange('G4:H' + lastRowInProjection);
            projectionValidationRange.setDataValidation(validation);
            
            Logger.log('  Updated projection tab: ' + projectionTab);
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
        
      } catch (e) {
        errors.push('Row ' + (i + 2) + ' (' + campus + '): Cannot open spreadsheet - ' + e.toString());
      }
    }
    
    // Prepare summary message
    var message = 'Campus files update complete.\n\n';
    
    if (updated.length > 0) {
      message += 'Successfully updated (' + updated.length + '):\n' + updated.join('\n') + '\n\n';
    }
    
    if (skipped.length > 0) {
      message += 'Skipped (' + skipped.length + '):\n' + skipped.join('\n') + '\n\n';
    }
    
    if (errors.length > 0) {
      message += 'Errors (' + errors.length + '):\n' + errors.join('\n');
    }
    
    if (updated.length === 0 && errors.length === 0) {
      message += 'No files were updated. Check that spreadsheet IDs exist in column E.';
    }
    
  Logger.log('FUNCTION END: updateCampusFilesWithSep25Updates - Updated ' + updated.length + ' campus files');
  Logger.log('SUMMARY: ' + message);
  
  } finally {
    // Always release the lock
    lock.releaseLock();
  }
}