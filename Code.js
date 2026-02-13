/**
 * Removes all data validation from column G (7th column) in all month sheets.
 * Run this once if you want to allow any value (including blanks) in column G.
 */
function removeValidationFromColumnsFandGAllMonths() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var months = getMonthNames_();
  months.forEach(function(month) {
    var sheet = ss.getSheetByName(month);
    if (!sheet) {
      Logger.log('Sheet not found: ' + month);
      return;
    }
    var lastRow = sheet.getMaxRows();
    // Remove validations from column G (7) and column F (6)
    var rangeG = sheet.getRange(1, 7, lastRow, 1); // column G, all rows
    var rangeF = sheet.getRange(1, 6, lastRow, 1); // column F, all rows
    Logger.log('Clearing validations in ' + month + ' G1:G' + lastRow);
    rangeG.clearDataValidations();
    Logger.log('Clearing validations in ' + month + ' F1:F' + lastRow);
    rangeF.clearDataValidations();
  });
  Logger.log('removeValidationFromColumnGAllMonths completed.');
}
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
    SpreadsheetApp.getUi().alert(
      "Another instance is already running. Please wait and try again."
    );
    return;
  }

  try {
    var timestamp = new Date().toISOString();
    Logger.log("FUNCTION START: createCampusSpreadsheets at " + timestamp);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var infoSheet = ss.getSheetByName("CampusBMCSheetInfo");
    if (!infoSheet) {
      throw new Error("CampusBMCSheetInfo sheet not found");
    }
    var lastRow = infoSheet.getLastRow();
    var data = infoSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var createdNames = [];
    var errorMessages = [];
    // Columns: A=email, B=campus, C=level, D=main/level folderId, E=campus spreadsheetId
    for (var i = 0; i < data.length; i++) {
      // start at row 2
      var row = data[i];
      var email = row[0];
      var campus = row[1];
      var folderId = row[3];
      var spreadsheetId = row[4];
      if (!campus) {
        errorMessages.push("Row " + (i + 2) + ": Missing campus name.");
        continue;
      }
      if (!folderId) {
        errorMessages.push(
          "Row " + (i + 2) + " (" + campus + "): Missing Main/Level Folder ID."
        );
        continue;
      }
      if (!email) {
        errorMessages.push(
          "Row " + (i + 2) + " (" + campus + "): Missing email."
        );
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
        errorMessages.push(
          "Row " + (i + 2) + " (" + campus + "): Invalid folder ID."
        );
        continue;
      }
      // Make a copy of the template (bound spreadsheet)
      var templateId = ss.getId();
      var campusName = campus + " BMC Class Count";
      Logger.log("CREATING SPREADSHEET: " + campusName + " for row " + (i + 2));
      var newFile = DriveApp.getFileById(templateId).makeCopy(
        campusName,
        folder
      );
      var newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
      // Remove CampusBMCSheetInfo and Totals sheets
      var sheetsToRemove = ["CampusBMCSheetInfo", "Totals"];
      sheetsToRemove.forEach(function (sheetName) {
        var sheet = newSpreadsheet.getSheetByName(sheetName);
        if (sheet) newSpreadsheet.deleteSheet(sheet);
      });
      // Write new spreadsheet ID in column E
      infoSheet.getRange(i + 2, 5).setValue(newFile.getId());
      // Share spreadsheet with email
      try {
        newFile.addEditor(email, false); // false = don't send notification
      } catch (e) {
        var errorMsg =
          "Row " +
          (i + 2) +
          " (" +
          campus +
          "): Could not share spreadsheet with " +
          email;
        if (
          e.toString().indexOf("permission") !== -1 ||
          e.toString().indexOf("sharing") !== -1
        ) {
          errorMsg += " (Check Shared Drive sharing permissions)";
        }
        errorMessages.push(errorMsg + ".");
      }
      createdNames.push(campusName);
    }
    var ui = SpreadsheetApp.getUi();
    var message = "";
    if (createdNames.length > 0) {
      message +=
        "Created " +
        createdNames.length +
        " spreadsheet(s):\n" +
        createdNames.join("\n") +
        "\n\n";
    } else {
      message += "No new spreadsheets were created.\n";
    }
    if (errorMessages.length > 0) {
      message += "Errors:\n" + errorMessages.join("\n");
    }

    Logger.log(
      "FUNCTION END: createCampusSpreadsheets - Created " +
        createdNames.length +
        " spreadsheets"
    );
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
        .addItem("Auto ES", "autoConsolidateLevelES")
        .addSeparator()
        .addItem("Start MS", "consolidateLevelStartMS")
        .addItem("Next Batch MS", "consolidateLevelNextBatchMS")
        .addItem("Auto MS", "autoConsolidateLevelMS")
        .addSeparator()
        .addItem("Start HS", "consolidateLevelStartHS")
        .addItem("Next Batch HS", "consolidateLevelNextBatchHS")
        .addItem("Auto HS", "autoConsolidateLevelHS")
        .addSeparator()
        .addItem("Auto All Levels", "autoConsolidateAllLevels")
        .addItem("Stop Auto Processing", "stopAutoProcessing")
        .addSeparator()
        .addItem("Show Status", "showConsolidationStatus")
    )
    .addSeparator()
    .addItem("Create Campus Spreadsheets", "createCampusSpreadsheets")
    .addToUi();
}

// ================= Consolidation by Level (ES/MS/HS) =================
// Public wrappers for menu
function consolidateLevelStartES() {
  consolidateLevelStart_("ES");
}
function consolidateLevelStartMS() {
  consolidateLevelStart_("MS");
}
function consolidateLevelStartHS() {
  consolidateLevelStart_("HS");
}
function consolidateLevelNextBatchES() {
  consolidateLevelNextBatch_("ES");
}
function consolidateLevelNextBatchMS() {
  consolidateLevelNextBatch_("MS");
}
function consolidateLevelNextBatchHS() {
  consolidateLevelNextBatch_("HS");
}

/**
 * Reset cursor for a level and process the first batch.
 * Clears previous data for this level's campuses only (preserving other levels' data).
 *
 * @param {string} level - The school level (ES|MS|HS).
 * @returns {void}
 */
function consolidateLevelStart_(level) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(levelCursorKey_(level), "0");
  // Reset the clear-once flag so a fresh run overwrites prior data for this level
  props.deleteProperty(levelClearedKey_(level));
  consolidateLevelNextBatch_(level);
}

/**
 * Process the next batch of campuses for the given level (ES/MS/HS).
 * Clears prior data for this level's campuses only (preserving other levels' data), then appends.
 * Reads campus data from row 3 and appends to existing data in the master.
 *
 * @param {string} level - The school level (ES|MS|HS).
 * @returns {void}
 */
function consolidateLevelNextBatch_(level) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "Another instance is running. Try again later."
    );
    return;
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var infoSheet = ss.getSheetByName("CampusBMCSheetInfo");
    if (!infoSheet) throw new Error("CampusBMCSheetInfo sheet not found");

    var months = getMonthNames_();
    var props = PropertiesService.getScriptProperties();
    var batchSize = parseInt(
      props.getProperty("CONSOLIDATE_BATCH_SIZE") || "15",
      10
    );
    var cursorKey = levelCursorKey_(level);
    var startIndex = parseInt(props.getProperty(cursorKey) || "0", 10);

    // Build list of campuses for this level
    var lastRow = infoSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("No rows in CampusBMCSheetInfo.");
      return;
    }
    var allRows = infoSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    // Columns: [A email, B campus, C level, D folderId, E spreadsheetId]
    var levelRows = [];
    for (var i = 0; i < allRows.length; i++) {
      var row = allRows[i];
      if ((row[2] || "").toString().trim().toUpperCase() !== level) continue;
      var campus = row[1];
      var ssId = row[4];
      if (!campus || !ssId) continue;
      levelRows.push({ campus: campus, id: ssId });
    }
    if (levelRows.length === 0) {
      SpreadsheetApp.getUi().alert(
        "No campuses found for level " + level + "."
      );
      return;
    }
    if (startIndex >= levelRows.length) {
      SpreadsheetApp.getUi().alert(
        "Level " + level + " already complete. Reset to run again."
      );
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
    batch.forEach(function (item) {
      var campusName = item.campus;
      var campusSs;
      try {
        campusSs = SpreadsheetApp.openById(item.id);
      } catch (e) {
        errors.push(
          "Skip " + campusName + ": cannot open spreadsheet " + item.id
        );
        return;
      }
      var foundAny = false;
      months.forEach(function (month) {
        var sh = campusSs.getSheetByName(month);
        if (!sh) return;
        var lr = sh.getLastRow();
        var lc = sh.getLastColumn();
        if (lr < 1 || lc < 1) return;

        // For projection sheets, data starts at row 4 (skip header at row 3)
        // For regular sheets, data starts at row 3
        var dataStartRow =
          month.toLowerCase().indexOf("projection") !== -1 ? 4 : 3;
        if (lr <= dataStartRow - 1) return; // headers/metadata only

        var values = sh
          .getRange(dataStartRow, 1, lr - dataStartRow + 1, lc)
          .getValues();
        // Only keep rows where at least one of columns A, B, C, or E:O is not empty (exclude D)
        var nonBlank = values.filter(function (r) {
          // r[0]=A, r[1]=B, r[2]=C, r[3]=D, r[4]=E, ..., r[14]=O
          for (var i = 0; i <= 2; i++) {
            if (r[i] && String(r[i]).trim() !== "") return true;
          }
          for (var i = 4; i <= 14; i++) {
            if (r[i] && String(r[i]).trim() !== "") return true;
          }
          return false;
        });
        if (nonBlank.length === 0) return;
        foundAny = true;
        // Prepare month bucket
        var bucket = byMonth[month];
        bucket.rowsByCampus[campusName] = (
          bucket.rowsByCampus[campusName] || []
        ).concat(nonBlank);
      });
      if (foundAny) processedCampuses.push(campusName);
    });

    // Clear existing data for this level once per run (only when starting from index 0)
    clearLevelDataOnce_(ss, months, level);

    // Write to master per month: append fresh rows (schema preserved by destination sheets)

    var appendSummary = [];
    var errorRows = [];
    var validValues = [
      "OHI",
      "ED",
      "AU",
      "SI",
      "CI",
      "ID",
      "LD",
      "OTHER",
      "Other",
      "2026-2027 Grade Level",
    ];
    var validValuesUpper = validValues.map(function (v) {
      return v.toUpperCase();
    });

    // Valid campus names from the error message + additional staff names found in data
    var validCampuses = [
      "Adams Hill",
      "Aue",
      "Behlau",
      "Bernal",
      "Blattman",
      "Boldt",
      "Boone",
      "Brandeis",
      "Brauchle",
      "Braun Station",
      "Brennan",
      "Briscoe",
      "Burke",
      "Cable",
      "Carlos Coon",
      "Carnahan",
      "Carson",
      "Chumbley",
      "Clark",
      "Cody",
      "Cole",
      "Colonies North",
      "Connally",
      "Driggers",
      "Ellison",
      "Elrod",
      "Evers",
      "Fields",
      "Fisher",
      "Folks",
      "Forester",
      "Franklin",
      "Galm",
      "Garcia",
      "Glass",
      "Glenn",
      "Glenoaks",
      "Harlan",
      "Hatchett",
      "Henderson",
      "Hobby",
      "Hoffmann",
      "Holmes",
      "Holmgreen MS",
      "Howsman",
      "Jay",
      "Jefferson",
      "Jones",
      "Jordan",
      "Kallison",
      "Krueger",
      "Kuentz",
      "Langley",
      "Lewis",
      "Lieck",
      "Linton",
      "Locke Hill",
      "Los Reyes",
      "Luna",
      "Marshall",
      "Martin",
      "Mary Hull",
      "May",
      "McDermott",
      "Mead",
      "Meadow Village",
      "Michael",
      "Mireles",
      "Murnin",
      "NAHS",
      "Neff",
      "Northside Alternative MS",
      "Northwest Crossing",
      "O'Connor",
      "Oak Hills Terrace",
      "Ott",
      "Passmore",
      "Pease",
      "Powell",
      "Raba",
      "Rawlinson",
      "Rayburn",
      "Reed",
      "Rhodes",
      "Ross",
      "Rudder",
      "Scarborough",
      "Scobee",
      "Sotomayor",
      "Steubing",
      "Stevens",
      "Stevenson",
      "Stinson",
      "Straus",
      "Taft",
      "Thornton",
      "Timberwilde",
      "Tomlinson",
      "Vale",
      "Valley Hi",
      "Wanke",
      "Ward",
      "Warren",
      "Wernli",
      "WWT",
      "Zachry",
      "Mora",
      "Nichols",
      "Fernandez",
    ];
    var validCampusesUpper = validCampuses.map(function (v) {
      return v.toUpperCase();
    });

    Logger.log(
      "Valid campuses array initialized with " + validCampuses.length + " items"
    );
    Logger.log(
      "Sample valid campuses (upper): " +
        validCampusesUpper.slice(0, 10).join(", ")
    );

    var ssIdMaster = ss.getId(); // to avoid accidental writes elsewhere
    var monthCount = 0;
    months.forEach(function (month) {
      // Add delay between months to prevent rate limiting (except for first month)
      if (monthCount > 0) {
        Logger.log("Adding delay before processing " + month + "...");
        Utilities.sleep(1000); // 1 second delay between months
      }
      monthCount++;

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
      campusNames.forEach(function (campusName) {
        var rows = bucket.rowsByCampus[campusName];
        rows.forEach(function (r, idx) {
          // Align to destination columns to avoid shifting columns with validations
          var aligned = r.slice(0, lcDest);
          while (aligned.length < lcDest) aligned.push("");

          // Clean and trim all string values to prevent validation errors
          for (var col = 0; col < aligned.length; col++) {
            if (aligned[col] && typeof aligned[col] === "string") {
              // Normalize common problematic characters: non-breaking spaces and smart quotes
              aligned[col] = aligned[col]
                .replace(/\u00A0/g, " ")
                .replace(/[\u2018\u2019]/g, "'")
                .trim();
            }
          }

          // Calculate actual row number based on sheet type
          var dataStartRow =
            month.toLowerCase().indexOf("projection") !== -1 ? 4 : 3;
          var actualRowNumber = idx + dataStartRow;

          // Check column F (index 5) - Service Type validation
          var valueF = aligned[5] ? String(aligned[5]).trim() : "";
          var valueFUpper = valueF.toUpperCase();
          if (valueF && validValuesUpper.indexOf(valueFUpper) === -1) {
            errorRows.push({
              campus: campusName,
              month: month,
              row: actualRowNumber,
              column: "F",
              value: aligned[5],
              error: "Invalid service type",
            });
            return; // skip this row
          }

          // Check column D (index 3) - Campus validation
          var valueD = aligned[3] ? String(aligned[3]).trim() : "";
          // Normalize apostrophes - replace curly/smart quotes with straight apostrophe
          valueD = valueD.replace(/[\u2018\u2019]/g, "'");
          var valueDUpper = valueD.toUpperCase();

          if (valueD && validCampusesUpper.indexOf(valueDUpper) === -1) {
            errorRows.push({
              campus: campusName,
              month: month,
              row: actualRowNumber,
              column: "D",
              value: aligned[3],
              error: "Invalid campus name",
            });
            return; // skip this row
          }

          rowsToAppend.push(aligned);
        });
      });

      if (rowsToAppend.length > 0) {
        // Final validation check before writing - catch any invalid values
        var finalValidationErrors = [];
        for (var i = 0; i < rowsToAppend.length; i++) {
          var row = rowsToAppend[i];
          var valueDFinal = row[3] ? String(row[3]).trim() : "";
          // Normalize apostrophes - replace curly/smart quotes with straight apostrophe
          valueDFinal = valueDFinal.replace(/[\u2018\u2019]/g, "'");
          var valueDFinalUpper = valueDFinal.toUpperCase();
          if (
            valueDFinal &&
            validCampusesUpper.indexOf(valueDFinalUpper) === -1
          ) {
            finalValidationErrors.push(
              "Row " +
                (i + 1) +
                ': "' +
                valueDFinal +
                '" (upper: "' +
                valueDFinalUpper +
                '")'
            );
            Logger.log(
              'INVALID VALUE FOUND: "' +
                valueDFinal +
                '" -> "' +
                valueDFinalUpper +
                '"'
            );
            Logger.log(
              "Valid campuses array contains: " +
                validCampusesUpper.slice(0, 10).join(", ") +
                "..."
            );
            Logger.log(
              'Index of "' +
                valueDFinalUpper +
                '" in valid array: ' +
                validCampusesUpper.indexOf(valueDFinalUpper)
            );
          }
        }

        if (finalValidationErrors.length > 0) {
          Logger.log(
            "FINAL VALIDATION FAILED for " +
              month +
              ": " +
              finalValidationErrors.join(", ")
          );
          appendSummary.push(
            month +
              ": SKIPPED - Contains invalid campus names: " +
              finalValidationErrors.join(", ")
          );
          return; // skip this entire month
        }

        // Log all values being written to column D for debugging
        var columnDValues = [];
        for (var i = 0; i < rowsToAppend.length; i++) {
          if (rowsToAppend[i][3]) {
            columnDValues.push('"' + rowsToAppend[i][3] + '"');
          }
        }
        if (columnDValues.length > 0) {
          Logger.log(
            "Writing to " + month + " column D: " + columnDValues.join(", ")
          );
        }

        // Append starting on row 3 in master sheets
        var startRow = Math.max(3, lrDest + 1);

        // Write data with retry logic and rate limiting protection
        var range = dest.getRange(startRow, 1, rowsToAppend.length, lcDest);
        var maxRetries = 3;
        var retryDelay = 2000; // 2 seconds
        var success = false;

        for (var attempt = 1; attempt <= maxRetries && !success; attempt++) {
          try {
            // Add delay to prevent rate limiting (except first attempt)
            if (attempt > 1) {
              Logger.log(
                "Attempt " +
                  attempt +
                  " for " +
                  month +
                  " after " +
                  retryDelay / 1000 +
                  " second delay..."
              );
              Utilities.sleep(retryDelay);
              retryDelay *= 2; // Exponential backoff
            }

            // Get current validation rules (only once, on first attempt)
            var dataValidations = [];
            if (attempt === 1) {
              try {
                // Batch get validation rules to reduce API calls
                var validationRange = dest.getRange(
                  startRow,
                  1,
                  rowsToAppend.length,
                  lcDest
                );
                var validations = validationRange.getDataValidations();
                dataValidations = validations;
              } catch (validationError) {
                Logger.log(
                  "Could not get validation rules, proceeding without them: " +
                    validationError.toString()
                );
                dataValidations = [];
              }
            }

            // Always attempt to clear validation rules before writing.
            // If we were able to read the prior validations we will restore them later.
            try {
              Logger.log(
                "Clearing data validations for target range: " +
                  startRow +
                  ",1 -> rows:" +
                  rowsToAppend.length +
                  ",cols:" +
                  lcDest
              );
              // Log a sample of the rows we'll write (first 5)
              Logger.log(
                "Sample rows to append: " +
                  JSON.stringify(rowsToAppend.slice(0, Math.min(5, rowsToAppend.length)))
              );
              range.clearDataValidations();
              Utilities.sleep(500); // Small delay after clearing validations
            } catch (clearError) {
              Logger.log(
                "Could not clear validation rules before writing: " +
                  clearError.toString()
              );
            }

            // Write the data
            range.setValues(rowsToAppend);

            // Small delay after writing data
            Utilities.sleep(500);

            // Restore validation rules if we had them
            if (dataValidations.length > 0) {
              try {
                range.setDataValidations(dataValidations);
              } catch (restoreError) {
                Logger.log(
                  "Could not restore validation rules: " +
                    restoreError.toString()
                );
                // Don't fail the whole operation if we can't restore validations
              }
            }

            success = true;
            appendSummary.push(month + ": " + rowsToAppend.length + " rows");
            Logger.log(
              "Successfully wrote " +
                rowsToAppend.length +
                " rows to " +
                month +
                " (attempt " +
                attempt +
                ")"
            );
          } catch (e) {
            Logger.log(
              "Attempt " +
                attempt +
                " failed for " +
                month +
                ": " +
                e.toString()
            );

            if (attempt === maxRetries) {
              // Last attempt failed, log detailed error info
              Logger.log(
                "All attempts failed for " + month + ". Error details:"
              );
              Logger.log("Error type: " + e.name);
              Logger.log("Error message: " + e.message);
              Logger.log("Rows to append: " + rowsToAppend.length);
              Logger.log(
                "Target range: " +
                  startRow +
                  ":" +
                  (startRow + rowsToAppend.length - 1)
              );
              Logger.log(
                "Sample row data: " + JSON.stringify(rowsToAppend.slice(0, 2))
              );

              // Check if this is a rate limiting or service error
              var errorMsg = e.toString().toLowerCase();
              if (
                errorMsg.indexOf("service") !== -1 ||
                errorMsg.indexOf("rate") !== -1 ||
                errorMsg.indexOf("quota") !== -1 ||
                errorMsg.indexOf("timeout") !== -1
              ) {
                // This looks like a temporary service issue
                throw new Error(
                  "Google Sheets service temporarily unavailable for " +
                    month +
                    ". Please try again in a few minutes. Original error: " +
                    e.toString()
                );
              } else {
                // Re-throw the original error
                throw e;
              }
            }
          }
        }
      }
    });

    // Advance cursor
    props.setProperty(cursorKey, String(endIndex));
    var done = endIndex >= levelRows.length;
    var errorMsg = "";
    if (errorRows.length > 0) {
      errorMsg = "VALIDATION ERRORS FOUND:\n";
      errorRows.forEach(function (e) {
        errorMsg +=
          "Campus: " +
          e.campus +
          ", Month: " +
          e.month +
          ", Row: " +
          e.row +
          ", Column: " +
          e.column +
          ', Value: "' +
          e.value +
          '", Error: ' +
          e.error +
          "\n";
      });
      errorMsg += "\n";
    }

    // Check if this was called as part of automated processing
    var isAutoMode = props.getProperty("AUTO_MODE_" + level) === "true";

    // Only show UI alert for manual processes
    if (!isAutoMode) {
      SpreadsheetApp.getUi().alert(
        errorMsg +
          "Level " +
          level +
          " batch complete.\nProcessed campuses: " +
          batch.length +
          "\n" +
          (appendSummary.length
            ? appendSummary.join(", ")
            : "No data this run") +
          "\nProgress: " +
          endIndex +
          " / " +
          levelRows.length +
          (done ? " (DONE)" : "")
      );
    } else {
      // For auto mode, just log the completion
      Logger.log(
        "Auto mode: Level " +
          level +
          " batch complete. Processed campuses: " +
          batch.length +
          ". Progress: " +
          endIndex +
          " / " +
          levelRows.length +
          (done ? " (DONE)" : "")
      );
    }

    if (isAutoMode && !done) {
      // Schedule the next batch (no UI in auto mode)
      Logger.log(
        "Auto mode: Scheduling next batch for " +
          level +
          ". Progress: " +
          endIndex +
          " / " +
          levelRows.length
      );
      scheduleNextBatch_(level, false);
    } else if (isAutoMode && done) {
      // This level is complete, clean up auto mode
      props.deleteProperty("AUTO_MODE_" + level);
      cleanupTriggers_(level);
      Logger.log("Auto mode: Level " + level + " completed successfully!");

      // Check if this was part of "Auto All Levels" processing
      var autoAllLevels = props.getProperty("AUTO_ALL_LEVELS") === "true";
      if (autoAllLevels) {
        startNextLevelInSequence_();
      } else {
        // Single level auto process complete - try to send email notification
        sendCompletionEmail_(level, "single");
      }
    }
  } catch (batchError) {
    Logger.log(
      "Error in consolidateLevelNextBatch_ for " +
        level +
        ": " +
        batchError.toString()
    );

    // Check if this is an automated process
    var isAutoMode = props.getProperty("AUTO_MODE_" + level) === "true";
    if (isAutoMode) {
      var errorMsg = batchError.toString().toLowerCase();
      if (
        errorMsg.indexOf("service") !== -1 ||
        errorMsg.indexOf("rate") !== -1 ||
        errorMsg.indexOf("quota") !== -1 ||
        errorMsg.indexOf("timeout") !== -1
      ) {
        // This looks like a temporary service issue, retry with longer delay
        Logger.log(
          "Service error detected, scheduling retry for " +
            level +
            " with longer delay..."
        );
        scheduleNextBatch_(level, true); // true = isRetry
        return; // Don't show error alert for temporary issues
      } else {
        // This is a more serious error, stop auto processing
        props.deleteProperty("AUTO_MODE_" + level);
        cleanupTriggers_(level);
        Logger.log(
          "Serious error in auto mode for " +
            level +
            ", stopping automation: " +
            batchError.toString()
        );
        sendErrorEmail_(level, batchError.toString());
        return; // Don't re-throw for auto mode
      }
    }

    // Re-throw the error for manual processes
    throw batchError;
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
  if (!n || n < 1) throw new Error("Invalid batch size");
  PropertiesService.getScriptProperties().setProperty(
    "CONSOLIDATE_BATCH_SIZE",
    String(n)
  );
}

// ---------------- helpers ----------------
/**
 * Property key for the per-level cursor position.
 * @param {string} level
 * @returns {string}
 */
function levelCursorKey_(level) {
  return "CONS_LEVEL_IDX_" + level.toUpperCase();
}
/**
 * Property key that marks whether a level's rows were cleared in the current run.
 * @param {string} level
 * @returns {string}
 */
function levelClearedKey_(level) {
  return "CONS_LEVEL_CLEARED_" + level.toUpperCase();
}

/**
 * Clears only data rows belonging to campuses of the specified level in each month sheet.
 * Preserves headers/metadata in rows 1–2 and data from other levels.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet.
 * @param {string[]} months - Month sheet names.
 * @param {string} level - ES|MS|HS.
 * @returns {void}
 */
function clearLevelDataOnce_(ss, months, level) {
  var props = PropertiesService.getScriptProperties();
  var key = levelClearedKey_(level);
  if (props.getProperty(key) === "true") return; // already cleared in this cycle

  // Get the list of campuses for this level
  var infoSheet = ss.getSheetByName("CampusBMCSheetInfo");
  if (!infoSheet) return;

  var lastRow = infoSheet.getLastRow();
  if (lastRow < 2) return;

  var allRows = infoSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var levelCampuses = [];

  // Build list of campus names for this level
  for (var i = 0; i < allRows.length; i++) {
    var row = allRows[i];
    if ((row[2] || "").toString().trim().toUpperCase() === level) {
      var campus = (row[1] || "").toString().trim();
      if (campus && levelCampuses.indexOf(campus) === -1) {
        levelCampuses.push(campus);
      }
    }
  }

  if (levelCampuses.length === 0) return;

  // For each month sheet, remove rows that contain data from level campuses
  months.forEach(function (month) {
    var sh = ss.getSheetByName(month);
    if (!sh) return;
    var lr = sh.getLastRow();
    var lc = sh.getLastColumn();

    // Determine data start row based on sheet name
    var dataStartRow = month === "APRIL/ MAY PROJECTIONS" ? 4 : 3;

    if (lr < dataStartRow || lc === 0) return; // No data rows to process

    // Read all data rows (starting from dataStartRow)
    var dataRange = sh.getRange(dataStartRow, 1, lr - dataStartRow + 1, lc);
    var allData = dataRange.getValues();

    // Filter out rows that belong to this level's campuses
    var rowsToKeep = [];
    for (var i = 0; i < allData.length; i++) {
      var row = allData[i];
      var campusInRow = "";

      // Column D (index 3) is the campus column based on validation logic
      if (row[3] && typeof row[3] === "string") {
        campusInRow = row[3].trim();
      }

      // Keep row if it doesn't belong to any campus in this level, or if it's completely empty
      var belongsToThisLevel = false;
      if (campusInRow) {
        for (var j = 0; j < levelCampuses.length; j++) {
          if (campusInRow.toLowerCase() === levelCampuses[j].toLowerCase()) {
            belongsToThisLevel = true;
            break;
          }
        }
      }

      if (!belongsToThisLevel) {
        // Clean up the row data before keeping it
        var cleanedRow = [];
        for (var k = 0; k < row.length; k++) {
          var cellValue = row[k];
          // Clean string values
          if (cellValue && typeof cellValue === "string") {
            cellValue = cellValue.trim();
            // For column K (index 10), ensure it matches expected validation values
            if (k === 10 && cellValue) {
              var lowerValue = cellValue.toLowerCase();
              if (lowerValue === "per day" || lowerValue === "perday") {
                cellValue = "per Day";
              } else if (
                lowerValue === "per week" ||
                lowerValue === "perweek"
              ) {
                cellValue = "per week";
              }
            }
          }
          cleanedRow.push(cellValue);
        }
        rowsToKeep.push(cleanedRow);
      }
    }

    // Clear all data rows first
    if (lr >= dataStartRow) {
      var clearRange = sh.getRange(dataStartRow, 1, lr - dataStartRow + 1, lc);

      // Clear data validation rules first to prevent validation errors during clearing
      try {
        clearRange.clearDataValidations();
        Utilities.sleep(200); // Small delay after clearing validations
      } catch (clearValidationError) {
        Logger.log(
          "Could not clear validation rules in " +
            month +
            ": " +
            clearValidationError.toString()
        );
      }
      // Then clear the content
      clearRange.clearContent();

      // Write back the rows we want to keep
      if (rowsToKeep.length > 0) {
        var keepRange = sh.getRange(dataStartRow, 1, rowsToKeep.length, lc);
        try {
          keepRange.setValues(rowsToKeep);
        } catch (writeError) {
          Logger.log(
            "Error writing preserved rows back to " +
              month +
              ": " +
              writeError.toString()
          );
          // If we can't write back the preserved rows, we need to restore the original data
          Logger.log("Attempting to restore original data for " + month);
          try {
            var restoreRange = sh.getRange(dataStartRow, 1, allData.length, lc);
            restoreRange.setValues(allData);
          } catch (restoreError) {
            Logger.log(
              "Failed to restore original data for " +
                month +
                ": " +
                restoreError.toString()
            );
          }
          throw writeError;
        }
      }
    }
  });

  props.setProperty(key, "true");
}

/**
 * Pads an array with empty strings until it reaches length n.
 * @param {any[]} arr - Row array.
 * @param {number} n - Desired length.
 * @returns {any[]} Padded copy.
 */
function padTo_(arr, n) {
  var a = arr.slice();
  while (a.length < n) a.push("");
  return a;
}

/**
 * Returns the list of month sheet names used by the template.
 * @returns {string[]} Month sheet names.
 */
function getMonthNames_() {
  return [
    "AUGUST",
    "SEPTEMBER",
    "OCTOBER",
    "NOVEMBER",
    "DECEMBER",
    "JANUARY",
    "FEBRUARY",
    "MARCH",
    "APRIL/ MAY PROJECTIONS",
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
    if (v !== null && v !== "" && String(v).trim() !== "") return false;
  }
  return true;
}

/**
 * Displays the current consolidation progress per level and batch size.
 * @returns {void}
 */
function showConsolidationStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var info = ss.getSheetByName("CampusBMCSheetInfo");
  if (!info) {
    SpreadsheetApp.getUi().alert("CampusBMCSheetInfo sheet not found.");
    return;
  }
  var lastRow = info.getLastRow();
  var rows = lastRow > 1 ? info.getRange(2, 1, lastRow - 1, 5).getValues() : [];
  // Count campuses with Spreadsheet ID per level
  var totals = { ES: 0, MS: 0, HS: 0 };
  rows.forEach(function (r) {
    var level = (r[2] || "").toString().trim().toUpperCase();
    var id = (r[4] || "").toString().trim();
    if (!id) return; // only count those with IDs
    if (totals.hasOwnProperty(level)) totals[level]++;
  });

  var props = PropertiesService.getScriptProperties();
  var batchSize = parseInt(
    props.getProperty("CONSOLIDATE_BATCH_SIZE") || "15",
    10
  );
  var idxES = parseInt(props.getProperty(levelCursorKey_("ES")) || "0", 10);
  var idxMS = parseInt(props.getProperty(levelCursorKey_("MS")) || "0", 10);
  var idxHS = parseInt(props.getProperty(levelCursorKey_("HS")) || "0", 10);

  var msg = [
    "Batch size: " + batchSize,
    "",
    "ES: " +
      Math.min(idxES, totals.ES) +
      " / " +
      totals.ES +
      (idxES >= totals.ES && totals.ES > 0 ? " (DONE)" : ""),
    "MS: " +
      Math.min(idxMS, totals.MS) +
      " / " +
      totals.MS +
      (idxMS >= totals.MS && totals.MS > 0 ? " (DONE)" : ""),
    "HS: " +
      Math.min(idxHS, totals.HS) +
      " / " +
      totals.HS +
      (idxHS >= totals.HS && totals.HS > 0 ? " (DONE)" : ""),
  ].join("\n");
  SpreadsheetApp.getUi().alert(msg);
}

// ================= Automated Batch Processing =================

/**
 * Start automated processing for ES level
 */
function autoConsolidateLevelES() {
  autoConsolidateLevel_("ES");
}

/**
 * Start automated processing for MS level
 */
function autoConsolidateLevelMS() {
  autoConsolidateLevel_("MS");
}

/**
 * Start automated processing for HS level
 */
function autoConsolidateLevelHS() {
  autoConsolidateLevel_("HS");
}

/**
 * Start automated processing for all levels in sequence (ES -> MS -> HS)
 */
function autoConsolidateAllLevels() {
  var props = PropertiesService.getScriptProperties();

  // Set up the sequence
  props.setProperty("AUTO_ALL_LEVELS", "true");
  props.setProperty("AUTO_SEQUENCE_CURRENT", "ES");
  props.setProperty("AUTO_SEQUENCE_REMAINING", "MS,HS");

  // Clean up any existing triggers
  stopAutoProcessing();

  // Start with ES
  SpreadsheetApp.getUi().alert(
    '🚀 Starting automated processing for all levels.\nSequence: ES → MS → HS\n\nProcessing will continue automatically overnight.\nUse "Stop Auto Processing" to cancel at any time.'
  );
  autoConsolidateLevel_("ES");
}

/**
 * Stop all automated processing and clean up triggers
 */
function stopAutoProcessing() {
  var props = PropertiesService.getScriptProperties();

  // Clear all auto mode flags
  props.deleteProperty("AUTO_MODE_ES");
  props.deleteProperty("AUTO_MODE_MS");
  props.deleteProperty("AUTO_MODE_HS");
  props.deleteProperty("AUTO_ALL_LEVELS");
  props.deleteProperty("AUTO_SEQUENCE_CURRENT");
  props.deleteProperty("AUTO_SEQUENCE_REMAINING");

  // Clean up all triggers
  cleanupTriggers_("ES");
  cleanupTriggers_("MS");
  cleanupTriggers_("HS");

  SpreadsheetApp.getUi().alert(
    "🛑 Automated processing stopped.\nAll triggers have been cleaned up."
  );
}

/**
 * Internal function to start automated processing for a specific level
 * @param {string} level - The school level (ES|MS|HS)
 */
function autoConsolidateLevel_(level) {
  var props = PropertiesService.getScriptProperties();

  // Check if already running
  if (props.getProperty("AUTO_MODE_" + level) === "true") {
    SpreadsheetApp.getUi().alert(
      "⚠️ Automated processing for " +
        level +
        ' is already running.\nUse "Stop Auto Processing" to cancel first.'
    );
    return;
  }

  // Clean up any existing triggers for this level
  cleanupTriggers_(level);

  // Set auto mode flag
  props.setProperty("AUTO_MODE_" + level, "true");

  // Initialize the level (reset cursor)
  consolidateLevelStart_(level);

  if (!props.getProperty("AUTO_ALL_LEVELS")) {
    SpreadsheetApp.getUi().alert(
      "🚀 Starting automated processing for " +
        level +
        ' level.\n\nProcessing will continue automatically overnight.\nUse "Stop Auto Processing" to cancel at any time.'
    );
  }
}

/**
 * Schedule the next batch for a level using time-based triggers
 * @param {string} level - The school level (ES|MS|HS)
 * @param {boolean} isRetry - Whether this is a retry after an error
 */
function scheduleNextBatch_(level, isRetry) {
  try {
    // Clean up any existing triggers for this level first
    cleanupTriggers_(level);

    // Use longer delay for retries to let service issues resolve
    var delayMinutes = isRetry ? 5 : 2; // 5 minutes for retries, 2 minutes normal

    // Create a new trigger to run after the delay
    var trigger = ScriptApp.newTrigger("autoContinueBatch_" + level)
      .timeBased()
      .after(delayMinutes * 60 * 1000) // Convert to milliseconds
      .create();

    // Store trigger ID for cleanup
    var props = PropertiesService.getScriptProperties();
    props.setProperty("TRIGGER_ID_" + level, trigger.getUniqueId());

    Logger.log(
      "Scheduled next batch for " +
        level +
        " in " +
        delayMinutes +
        " minutes. Trigger ID: " +
        trigger.getUniqueId()
    );
  } catch (error) {
    Logger.log(
      "Error scheduling next batch for " + level + ": " + error.toString()
    );
    // Fall back to longer immediate processing delay if trigger creation fails
    var fallbackDelay = isRetry ? 10000 : 5000; // 10 seconds for retries, 5 seconds normal
    Utilities.sleep(fallbackDelay);
    consolidateLevelNextBatch_(level);
  }
}

/**
 * Clean up triggers for a specific level
 * @param {string} level - The school level (ES|MS|HS)
 */
function cleanupTriggers_(level) {
  var props = PropertiesService.getScriptProperties();
  var triggerId = props.getProperty("TRIGGER_ID_" + level);

  if (triggerId) {
    try {
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getUniqueId() === triggerId) {
          ScriptApp.deleteTrigger(triggers[i]);
          Logger.log("Cleaned up trigger for " + level + ": " + triggerId);
          break;
        }
      }
    } catch (error) {
      Logger.log(
        "Error cleaning up trigger for " + level + ": " + error.toString()
      );
    }
    props.deleteProperty("TRIGGER_ID_" + level);
  }

  // Also clean up any orphaned triggers with our function names
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handlerFunction = triggers[i].getHandlerFunction();
    if (handlerFunction === "autoContinueBatch_" + level) {
      try {
        ScriptApp.deleteTrigger(triggers[i]);
        Logger.log(
          "Cleaned up orphaned trigger for " +
            level +
            ": " +
            triggers[i].getUniqueId()
        );
      } catch (error) {
        Logger.log("Error cleaning up orphaned trigger: " + error.toString());
      }
    }
  }
}

/**
 * Continue to the next level in "Auto All Levels" sequence
 */
function startNextLevelInSequence_() {
  var props = PropertiesService.getScriptProperties();
  var remaining = props.getProperty("AUTO_SEQUENCE_REMAINING");

  if (!remaining) {
    // All levels complete
    props.deleteProperty("AUTO_ALL_LEVELS");
    props.deleteProperty("AUTO_SEQUENCE_CURRENT");
    props.deleteProperty("AUTO_SEQUENCE_REMAINING");
    Logger.log("🎉 All levels completed successfully! ES ✅ MS ✅ HS ✅");
    sendCompletionEmail_("ALL", "sequence");
    return;
  }

  var levels = remaining.split(",");
  var nextLevel = levels.shift();
  var newRemaining = levels.join(",");

  // Update sequence tracking
  props.setProperty("AUTO_SEQUENCE_CURRENT", nextLevel);
  if (newRemaining) {
    props.setProperty("AUTO_SEQUENCE_REMAINING", newRemaining);
  } else {
    props.deleteProperty("AUTO_SEQUENCE_REMAINING");
  }

  // Start the next level
  Logger.log("Starting next level in sequence: " + nextLevel);
  autoConsolidateLevel_(nextLevel);
}

// ================= Auto-Continue Functions (called by triggers) =================

/**
 * Auto-continue function for ES (called by time-based triggers)
 */
function autoContinueBatch_ES() {
  autoContinueBatch_("ES");
}

/**
 * Auto-continue function for MS (called by time-based triggers)
 */
function autoContinueBatch_MS() {
  autoContinueBatch_("MS");
}

/**
 * Auto-continue function for HS (called by time-based triggers)
 */
function autoContinueBatch_HS() {
  autoContinueBatch_("HS");
}

/**
 * Internal function to continue batch processing
 * @param {string} level - The school level (ES|MS|HS)
 */
function autoContinueBatch_(level) {
  var props = PropertiesService.getScriptProperties();

  // Check if auto mode is still active
  if (props.getProperty("AUTO_MODE_" + level) !== "true") {
    Logger.log("Auto mode no longer active for " + level + ", stopping.");
    cleanupTriggers_(level);
    return;
  }

  try {
    // Clean up the trigger that called this function
    cleanupTriggers_(level);

    // Process the next batch
    consolidateLevelNextBatch_(level);
  } catch (error) {
    Logger.log("Error in auto-continue for " + level + ": " + error.toString());

    // Clean up auto mode on error
    props.deleteProperty("AUTO_MODE_" + level);
    cleanupTriggers_(level);

    // Send error notification instead of showing UI alert
    sendErrorEmail_(level, error.toString());
  }
}

// ================= Email Notification Functions =================

/**
 * Send completion email notification
 * @param {string} level - The school level or 'ALL'
 * @param {string} type - 'single' or 'sequence'
 */
function sendCompletionEmail_(level, type) {
  try {
    var user = Session.getActiveUser().getEmail();
    if (!user) {
      Logger.log("No active user email found, skipping completion email");
      return;
    }

    var subject, body;
    var spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
    var spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    if (type === "sequence") {
      subject = "🎉 BMC Consolidation Complete - All Levels";
      body =
        "Great news! The automated BMC consolidation process has successfully completed all levels.\n\n" +
        "✅ ES Level: Complete\n" +
        "✅ MS Level: Complete\n" +
        "✅ HS Level: Complete\n\n" +
        "Your consolidated report is ready for review.\n\n" +
        "Spreadsheet: " +
        spreadsheetName +
        "\n" +
        "View Report: " +
        spreadsheetUrl +
        "\n\n" +
        "Generated automatically by BMC Class Count Template Project";
    } else {
      subject = "✅ BMC Consolidation Complete - " + level + " Level";
      body =
        "The automated BMC consolidation process has successfully completed for the " +
        level +
        " level.\n\n" +
        "Your consolidated data is ready for review.\n\n" +
        "Spreadsheet: " +
        spreadsheetName +
        "\n" +
        "View Report: " +
        spreadsheetUrl +
        "\n\n" +
        "Generated automatically by BMC Class Count Template Project";
    }

    MailApp.sendEmail(user, subject, body);
    Logger.log("Completion email sent to: " + user);
  } catch (emailError) {
    Logger.log("Could not send completion email: " + emailError.toString());
  }
}

/**
 * Send error email notification
 * @param {string} level - The school level
 * @param {string} errorMessage - The error message
 */
function sendErrorEmail_(level, errorMessage) {
  try {
    var user = Session.getActiveUser().getEmail();
    if (!user) {
      Logger.log("No active user email found, skipping error email");
      return;
    }

    var spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
    var spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    var subject = "❌ BMC Consolidation Error - " + level + " Level";
    var body =
      "The automated BMC consolidation process encountered an error and has been stopped.\n\n" +
      "Level: " +
      level +
      "\n" +
      "Error: " +
      errorMessage +
      "\n\n" +
      'You can resume the process manually using the "Next Batch ' +
      level +
      '" option in the BMC menu.\n\n' +
      "Spreadsheet: " +
      spreadsheetName +
      "\n" +
      "Access Spreadsheet: " +
      spreadsheetUrl +
      "\n\n" +
      "Generated automatically by BMC Class Count Template Project";

    MailApp.sendEmail(user, subject, body);
    Logger.log("Error email sent to: " + user);
  } catch (emailError) {
    Logger.log("Could not send error email: " + emailError.toString());
  }
}
