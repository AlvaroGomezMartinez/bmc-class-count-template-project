/**
 * Creates a separate spreadsheet for each campus listed in CampusBMCSheetInfo.
 * - Removes CampusBMCSheetInfo and Totals from each copy
 * - Moves the copy to the folder in column D
 * - Shares with the email in column A
 * - Writes the new spreadsheet ID in column E
 *
 * @returns {void}
 */
function createCampusSpreadsheets() {
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
    Logger.log('FUNCTION START: createCampusSpreadsheets at ' + timestamp);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var infoSheet = ss.getSheetByName('CampusBMCSheetInfo');
    if (!infoSheet) {
      throw new Error('CampusBMCSheetInfo sheet not found');
    }
  var lastRow = infoSheet.getLastRow();
  var data = infoSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var createdNames = [];
  var errorMessages = [];
  // Columns: A=email, B=campus, C=level, D=main/level folderId, E=campus spreadsheetId
  for (var i = 0; i < data.length; i++) { // start at row 2
    var row = data[i];
    var email = row[0];
    var campus = row[1];
    var folderId = row[3];
    var spreadsheetId = row[4];
    if (!campus) {
      errorMessages.push('Row ' + (i+2) + ': Missing campus name.');
      continue;
    }
    if (!folderId) {
      errorMessages.push('Row ' + (i+2) + ' (' + campus + '): Missing Main/Level Folder ID.');
      continue;
    }
    if (!email) {
      errorMessages.push('Row ' + (i+2) + ' (' + campus + '): Missing email.');
      continue;
    }
    var fileExists = false;
    if (spreadsheetId) {
      try {
        var file = DriveApp.getFileById(spreadsheetId);
        file.getName(); // Will throw if file doesn't exist
        fileExists = true;
      } catch (e) {
        fileExists = false;
      }
    }
    if (fileExists) continue; // skip if file exists
    // Validate folder
    var folder;
    try {
      folder = DriveApp.getFolderById(folderId);
      folder.getName(); // Will throw if folder doesn't exist
    } catch (e) {
      errorMessages.push('Row ' + (i+2) + ' (' + campus + '): Invalid folder ID.');
      continue;
    }
    // Make a copy of the template (bound spreadsheet)
    var templateId = ss.getId();
    var campusName = campus + ' BMC Class Count';
    Logger.log('CREATING SPREADSHEET: ' + campusName + ' for row ' + (i+2));
    var newFile = DriveApp.getFileById(templateId).makeCopy(campusName, folder);
    var newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
    // Remove CampusBMCSheetInfo and Totals sheets
    var sheetsToRemove = ['CampusBMCSheetInfo', 'Totals'];
    sheetsToRemove.forEach(function(sheetName) {
      var sheet = newSpreadsheet.getSheetByName(sheetName);
      if (sheet) newSpreadsheet.deleteSheet(sheet);
    });
    // Write new spreadsheet ID in column E
    infoSheet.getRange(i+2, 5).setValue(newFile.getId());
    // Share spreadsheet with email
    try {
      newFile.addEditor(email, false); // false = don't send notification
    } catch (e) {
      var errorMsg = 'Row ' + (i+2) + ' (' + campus + '): Could not share spreadsheet with ' + email;
      if (e.toString().indexOf('permission') !== -1 || e.toString().indexOf('sharing') !== -1) {
        errorMsg += ' (Check Shared Drive sharing permissions)';
      }
      errorMessages.push(errorMsg + '.');
    }
    createdNames.push(campusName);
  }
  var ui = SpreadsheetApp.getUi();
  var message = '';
  if (createdNames.length > 0) {
    message += 'Created ' + createdNames.length + ' spreadsheet(s):\n' + createdNames.join('\n') + '\n\n';
  } else {
    message += 'No new spreadsheets were created.\n';
  }
  if (errorMessages.length > 0) {
    message += 'Errors:\n' + errorMessages.join('\n');
  }
  
  Logger.log('FUNCTION END: createCampusSpreadsheets - Created ' + createdNames.length + ' spreadsheets');
  ui.alert(message);
  } finally {
    // Always release the lock
    lock.releaseLock();
  }
}

/**
 * Test function for createCampusSpreadsheets logic using mock data.
 * Logs results to help verify correct behavior.
 *
 * @returns {void}
 */
function testCreateCampusSpreadsheets() {
  // Mock data: [email, campus, level, folderId, spreadsheetId]
  var mockData = [
    ['user1@example.com', 'North Campus', 'HS', 'FOLDERID1', ''],
    ['user2@example.com', 'South Campus', 'MS', 'FOLDERID2', 'EXISTINGID'],
    ['', 'East Campus', 'ES', 'FOLDERID3', ''],
    ['user4@example.com', '', 'HS', 'FOLDERID4', ''],
    ['user5@example.com', 'West Campus', 'HS', '', ''],
    ['user6@example.com', 'Central Campus', 'HS', 'BADFOLDERID', ''],
  ];
  var createdNames = [];
  var errorMessages = [];
  for (var i = 0; i < mockData.length; i++) {
    var row = mockData[i];
    var email = row[0];
    var campus = row[1];
    var folderId = row[3];
    var spreadsheetId = row[4];
    if (!campus) {
      errorMessages.push('Row ' + (i+2) + ': Missing campus name.');
      continue;
    }
    if (!folderId) {
      errorMessages.push('Row ' + (i+2) + ' (' + campus + '): Missing Main/Level Folder ID.');
      continue;
    }
    if (!email) {
      errorMessages.push('Row ' + (i+2) + ' (' + campus + '): Missing email.');
      continue;
    }
    var fileExists = false;
    if (spreadsheetId) {
      // Simulate file existence: only 'EXISTINGID' exists
      if (spreadsheetId === 'EXISTINGID') fileExists = true;
    }
    if (fileExists) continue; // skip if file exists
    // Validate folder: only FOLDERID1 and FOLDERID2 are valid
    if (folderId !== 'FOLDERID1' && folderId !== 'FOLDERID2') {
      errorMessages.push('Row ' + (i+2) + ' (' + campus + '): Invalid folder ID.');
      continue;
    }
    // Simulate making a copy and removing sheets
    var campusName = campus + ' BMC Class Count';
    // Simulate removal of CampusBMCSheetInfo and Totals
    var sheets = ['Sheet1', 'CampusBMCSheetInfo', 'Totals', 'OtherSheet'];
    sheets = sheets.filter(function(name) {
      return name !== 'CampusBMCSheetInfo' && name !== 'Totals';
    });
    if (sheets.indexOf('Totals') !== -1) {
      errorMessages.push('Row ' + (i+2) + ' (' + campus + '): Totals sheet was not removed.');
    }
    createdNames.push(campusName);
  }
  Logger.log('Created: ' + createdNames.join(', '));
  Logger.log('Errors: ' + errorMessages.join(', '));
}

/**
 * Adds custom menu items to the spreadsheet UI for consolidation and setup actions.
 * @returns {void}
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚩 BMC")
    .addSubMenu(
      ui
        .createMenu("Get Campus Data")
        .addItem("Start ES", "consolidateLevelStartES")
        .addItem("Next Batch ES", "consolidateLevelNextBatchES")
        .addSeparator()
        .addItem("Start MS", "consolidateLevelStartMS")
        .addItem("Next Batch MS", "consolidateLevelNextBatchMS")
        .addSeparator()
        .addItem("Start HS", "consolidateLevelStartHS")
        .addItem("Next Batch HS", "consolidateLevelNextBatchHS")
        .addSeparator()
  .addItem("Show Status", "showConsolidationStatus")
    )
    .addSeparator()
    .addItem("Create Campus Spreadsheets", "createCampusSpreadsheets")
    .addToUi();
}

// ================= Consolidation by Level (ES/MS/HS) =================
// Public wrappers for menu
function consolidateLevelStartES() { consolidateLevelStart_('ES'); }
function consolidateLevelStartMS() { consolidateLevelStart_('MS'); }
function consolidateLevelStartHS() { consolidateLevelStart_('HS'); }
function consolidateLevelNextBatchES() { consolidateLevelNextBatch_('ES'); }
function consolidateLevelNextBatchMS() { consolidateLevelNextBatch_('MS'); }
function consolidateLevelNextBatchHS() { consolidateLevelNextBatch_('HS'); }

/**
 * Reset cursor for a level and process the first batch.
 * Clears previous level data (rows 3+) once at the start of a run.
 *
 * @param {string} level - The school level (ES|MS|HS).
 * @returns {void}
 */
function consolidateLevelStart_(level) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(levelCursorKey_(level), '0');
  // Reset the clear-once flag so a fresh run overwrites prior data for this level
  props.deleteProperty(levelClearedKey_(level));
  consolidateLevelNextBatch_(level);
}

/**
 * Process the next batch of campuses for the given level (ES/MS/HS).
 * Overwrites prior data for those campuses by clearing level rows once per run, then appending.
 * Reads campus data from row 3 and appends to row 3 in the master.
 *
 * @param {string} level - The school level (ES|MS|HS).
 * @returns {void}
 */
function consolidateLevelNextBatch_(level) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Another instance is running. Try again later.');
    return;
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var infoSheet = ss.getSheetByName('CampusBMCSheetInfo');
    if (!infoSheet) throw new Error('CampusBMCSheetInfo sheet not found');

    var months = getMonthNames_();
    var props = PropertiesService.getScriptProperties();
  var batchSize = parseInt(props.getProperty('CONSOLIDATE_BATCH_SIZE') || '15', 10);
    var cursorKey = levelCursorKey_(level);
    var startIndex = parseInt(props.getProperty(cursorKey) || '0', 10);

    // Build list of campuses for this level
    var lastRow = infoSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No rows in CampusBMCSheetInfo.');
      return;
    }
    var allRows = infoSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    // Columns: [A email, B campus, C level, D folderId, E spreadsheetId]
    var levelRows = [];
    for (var i = 0; i < allRows.length; i++) {
      var row = allRows[i];
      if ((row[2] || '').toString().trim().toUpperCase() !== level) continue;
      var campus = row[1];
      var ssId = row[4];
      if (!campus || !ssId) continue;
      levelRows.push({ campus: campus, id: ssId });
    }
    if (levelRows.length === 0) {
      SpreadsheetApp.getUi().alert('No campuses found for level ' + level + '.');
      return;
    }
    if (startIndex >= levelRows.length) {
      SpreadsheetApp.getUi().alert('Level ' + level + ' already complete. Reset to run again.');
      return;
    }

    var endIndex = Math.min(levelRows.length, startIndex + batchSize);
    var batch = levelRows.slice(startIndex, endIndex);

    // For each month, collect new rows per campus (no headers; preserve dest sheet schema)
    var byMonth = {}; // month -> { rowsByCampus: { campus -> rows[][] } }
    for (var m = 0; m < months.length; m++) {
      byMonth[months[m]] = { rowsByCampus: {} };
    }

    var errors = [];
    var processedCampuses = [];
    batch.forEach(function(item) {
      var campusName = item.campus;
      var campusSs;
      try {
        campusSs = SpreadsheetApp.openById(item.id);
      } catch (e) {
        errors.push('Skip ' + campusName + ': cannot open spreadsheet ' + item.id);
        return;
      }
      var foundAny = false;
      months.forEach(function(month) {
        var sh = campusSs.getSheetByName(month);
        if (!sh) return;
        var lr = sh.getLastRow();
        var lc = sh.getLastColumn();
        if (lr < 1 || lc < 1) return;
  // Data begins on row 3 in campus sheets
  if (lr <= 2) return; // headers/metadata only
  var values = sh.getRange(3, 1, lr - 2, lc).getValues();
        // Filter empty rows
        var nonBlank = values.filter(function(r) { return !isRowEmpty_(r); });
        if (nonBlank.length === 0) return;
        foundAny = true;
        // Prepare month bucket
        var bucket = byMonth[month];
        bucket.rowsByCampus[campusName] = (bucket.rowsByCampus[campusName] || []).concat(nonBlank);
      });
      if (foundAny) processedCampuses.push(campusName);
    });

    // Clear existing data for this level once per run (only when starting from index 0)
    clearLevelDataOnce_(ss, months, level);

    // Write to master per month: append fresh rows (schema preserved by destination sheets)
    var appendSummary = [];
    var ssIdMaster = ss.getId(); // to avoid accidental writes elsewhere
    months.forEach(function(month) {
      var bucket = byMonth[month];
      var campusNames = Object.keys(bucket.rowsByCampus);
      if (campusNames.length === 0) return;
      var dest = ss.getSheetByName(month);
      if (!dest) {
        // Month sheet must exist in the master to preserve validations; if not, create and skip validations.
        dest = ss.insertSheet(month);
      }

  var lrDest = dest.getLastRow();
      var lcDest = dest.getLastColumn();
      if (lcDest === 0) {
        // If empty, try to adopt the first campus row width
        var firstCampus = campusNames[0];
        var firstRows = bucket.rowsByCampus[firstCampus];
        lcDest = firstRows && firstRows[0] ? firstRows[0].length : 1;
      }

      var rowsToAppend = [];
      campusNames.forEach(function(campusName) {
        var rows = bucket.rowsByCampus[campusName];
        rows.forEach(function(r) {
          // Align to destination columns to avoid shifting columns with validations
          var aligned = r.slice(0, lcDest);
          while (aligned.length < lcDest) aligned.push('');
          rowsToAppend.push(aligned);
        });
      });

      if (rowsToAppend.length > 0) {
        // Append starting on row 3 in master sheets
        var startRow = Math.max(3, lrDest + 1);
        dest.getRange(startRow, 1, rowsToAppend.length, lcDest).setValues(rowsToAppend);
        appendSummary.push(month + ': ' + rowsToAppend.length + ' rows');
      }
    });

    // Advance cursor
    props.setProperty(cursorKey, String(endIndex));
    var done = endIndex >= levelRows.length;
    SpreadsheetApp.getUi().alert('Level ' + level + ' batch complete.\nProcessed campuses: ' + batch.length + '\n' + (appendSummary.length ? appendSummary.join(', ') : 'No data this run') + '\nProgress: ' + endIndex + ' / ' + levelRows.length + (done ? ' (DONE)' : ''));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Optionally set the batch size globally (default 15 if unset).
 *
 * @param {number|string} size - Number of campuses per batch (>=1).
 * @returns {void}
 */
function setConsolidationBatchSize(size) {
  var n = parseInt(size, 10);
  if (!n || n < 1) throw new Error('Invalid batch size');
  PropertiesService.getScriptProperties().setProperty('CONSOLIDATE_BATCH_SIZE', String(n));
}

// ---------------- helpers ----------------
/**
 * Property key for the per-level cursor position.
 * @param {string} level
 * @returns {string}
 */
function levelCursorKey_(level) { return 'CONS_LEVEL_IDX_' + level.toUpperCase(); }
/**
 * Property key that marks whether a level's rows were cleared in the current run.
 * @param {string} level
 * @returns {string}
 */
function levelClearedKey_(level) { return 'CONS_LEVEL_CLEARED_' + level.toUpperCase(); }

/**
 * Clears all data rows (row 3+) in each month sheet for a level exactly once per Start cycle.
 * Preserves headers/metadata in rows 1–2.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet.
 * @param {string[]} months - Month sheet names.
 * @param {string} level - ES|MS|HS.
 * @returns {void}
 */
function clearLevelDataOnce_(ss, months, level) {
  var props = PropertiesService.getScriptProperties();
  var key = levelClearedKey_(level);
  if (props.getProperty(key) === 'true') return; // already cleared in this cycle
  months.forEach(function(month) {
    var sh = ss.getSheetByName(month);
    if (!sh) return;
    var lr = sh.getLastRow();
    var lc = sh.getLastColumn();
    // Clear rows 3+ only to preserve headers/metadata in rows 1–2
    if (lr > 2 && lc > 0) {
      sh.getRange(3, 1, lr - 2, lc).clearContent();
    }
  });
  props.setProperty(key, 'true');
}

/**
 * Pads an array with empty strings until it reaches length n.
 * @param {any[]} arr - Row array.
 * @param {number} n - Desired length.
 * @returns {any[]} Padded copy.
 */
function padTo_(arr, n) {
  var a = arr.slice();
  while (a.length < n) a.push('');
  return a;
}

/**
 * Returns the list of month sheet names used by the template.
 * @returns {string[]} Month sheet names.
 */
function getMonthNames_() {
  return [
    'AUGUST',
    'SEPTEMBER',
    'OCTOBER',
    'NOVEMBER',
    'DECEMBER',
    'JANUARY',
    'FEBRUARY',
    'MARCH',
    'APRIL/MAY PROJECTIONS'
  ];
}

/**
 * Checks if a row is effectively empty (all cells null/empty/whitespace).
 * @param {any[]} row
 * @returns {boolean}
 */
function isRowEmpty_(row) {
  for (var i = 0; i < row.length; i++) {
    var v = row[i];
    if (v !== null && v !== '' && String(v).trim() !== '') return false;
  }
  return true;
}

/**
 * Displays the current consolidation progress per level and batch size.
 * @returns {void}
 */
function showConsolidationStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var info = ss.getSheetByName('CampusBMCSheetInfo');
  if (!info) {
    SpreadsheetApp.getUi().alert('CampusBMCSheetInfo sheet not found.');
    return;
  }
  var lastRow = info.getLastRow();
  var rows = lastRow > 1 ? info.getRange(2, 1, lastRow - 1, 5).getValues() : [];
  // Count campuses with Spreadsheet ID per level
  var totals = { ES: 0, MS: 0, HS: 0 };
  rows.forEach(function(r){
    var level = (r[2] || '').toString().trim().toUpperCase();
    var id = (r[4] || '').toString().trim();
    if (!id) return; // only count those with IDs
    if (totals.hasOwnProperty(level)) totals[level]++;
  });

  var props = PropertiesService.getScriptProperties();
  var batchSize = parseInt(props.getProperty('CONSOLIDATE_BATCH_SIZE') || '15', 10);
  var idxES = parseInt(props.getProperty(levelCursorKey_('ES')) || '0', 10);
  var idxMS = parseInt(props.getProperty(levelCursorKey_('MS')) || '0', 10);
  var idxHS = parseInt(props.getProperty(levelCursorKey_('HS')) || '0', 10);

  var msg = [
    'Batch size: ' + batchSize,
    '',
    'ES: ' + Math.min(idxES, totals.ES) + ' / ' + totals.ES + (idxES >= totals.ES && totals.ES > 0 ? ' (DONE)' : ''),
    'MS: ' + Math.min(idxMS, totals.MS) + ' / ' + totals.MS + (idxMS >= totals.MS && totals.MS > 0 ? ' (DONE)' : ''),
    'HS: ' + Math.min(idxHS, totals.HS) + ' / ' + totals.HS + (idxHS >= totals.HS && totals.HS > 0 ? ' (DONE)' : '')
  ].join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Copies the data validation rule from AUGUST!D3:D997 in the main spreadsheet
 * and applies it to D3:D (down to last row) in every month sheet of every campus spreadsheet listed in CampusBMCInfo!E.
 *
 * @returns {void}
 */
function copyValidationToAllMonthsAllCampuses() {
  var MAIN_SPREADSHEET_ID = '1iIkKYUMsc7Lo8CZXBryOBRccIFtMcOdJP4aANeKejgs';
  var MONTHS = getMonthNames_();

  // Get validation rule from main spreadsheet AUGUST!D3:D997
  var mainSs = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  var augustSheet = mainSs.getSheetByName('AUGUST');
  if (!augustSheet) {
    SpreadsheetApp.getUi().alert('AUGUST sheet not found in main spreadsheet.');
    return;
  }
  var validationRules = augustSheet.getRange(3, 4, 995, 1).getDataValidations(); // D3:D997

  // Get campus spreadsheet IDs from CampusBMCInfo!E
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var infoSheet = ss.getSheetByName('CampusBMCInfo');
  if (!infoSheet) {
    SpreadsheetApp.getUi().alert('CampusBMCInfo sheet not found.');
    return;
  }
  var lastRow = infoSheet.getLastRow();
  var ids = infoSheet.getRange(2, 5, lastRow - 1, 1).getValues()
    .map(function(row) { return row[0]; })
    .filter(function(id) { return id && id.toString().trim() !== ''; });

  var errors = [];
  var updated = 0;
  ids.forEach(function(id) {
    try {
      var campusSs = SpreadsheetApp.openById(id);
      MONTHS.forEach(function(month) {
        var sheet = campusSs.getSheetByName(month);
        if (!sheet) return;
        var lr = sheet.getLastRow();
        if (lr < 3) return; // No data rows
        var nRows = lr - 2;
        var range = sheet.getRange(3, 4, nRows, 1); // D3:D(lastRow)
        // Truncate or repeat validation rules as needed
        var rulesToApply = [];
        for (var i = 0; i < nRows; i++) {
          rulesToApply.push(validationRules[i % validationRules.length][0]);
        }
        range.setDataValidations(rulesToApply.map(function(rule){ return [rule]; }));
        updated++;
      });
    } catch (e) {
      errors.push('Could not update spreadsheet ID ' + id + ': ' + e);
    }
  });

  var msg = 'Validation copied to ' + updated + ' sheet(s).';
  if (errors.length) msg += '\nErrors:\n' + errors.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

