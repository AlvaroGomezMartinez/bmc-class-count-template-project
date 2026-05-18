[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

> Originally created for Northside Independent School District (NISD). Generalized for public use upon retirement.

# BMC Class Count Template Project

A Google Apps Script bound to a master Google Sheets spreadsheet that automates the consolidation of monthly data from individual campus spreadsheets into a centralized dashboard. Data is collected in resumable batches by school level (ES / MS / HS).

## What it does

- **Creates campus spreadsheets** — copies the master template once per campus, moves each copy to the correct Drive folder, shares it with the campus editor, and records the new spreadsheet ID in the control sheet.
- **Consolidates monthly data** — reads each campus spreadsheet's monthly tabs and appends the data into the corresponding monthly tabs of the master spreadsheet, organized by school level.
- **Preserves schema** — headers and data validation rules in the master's monthly tabs are never modified; incoming rows are trimmed or padded to match the master's column count.
- **Resumable batching** — processes campuses in configurable batches (default 15) using a cursor stored in Script Properties, so long runs can be paused and resumed without data loss.
- **Automated overnight processing** — optional time-based triggers run batches automatically and send email notifications on completion or error.

## Spreadsheet structure

### Control sheet: `CampusBMCSheetInfo`

| Column | Contents |
|--------|----------|
| A | Editor email — the campus user the spreadsheet is shared with |
| B | Campus name |
| C | School level (`ES`, `MS`, or `HS`) |
| D | Google Drive folder ID where the campus spreadsheet should be stored |
| E | Campus spreadsheet ID — populated automatically when the spreadsheet is created |

### Monthly tabs (in each campus spreadsheet and the master)

`AUGUST`, `SEPTEMBER`, `OCTOBER`, `NOVEMBER`, `DECEMBER`, `JANUARY`, `FEBRUARY`, `MARCH`, `APRIL/ MAY PROJECTIONS`

- August–March: data starts on row 3 (rows 1–2 are headers/metadata/validation).
- April/May Projections: data starts on row 4 (rows 1–3 are headers/metadata/validation).

## Setup

### 1. Clone the repo and push to Apps Script

```bash
git clone <repo-url>
cd bmc-class-count-template-project
clasp push
```

> Requires [clasp](https://github.com/google/clasp) and a `.clasp.json` file with your `scriptId`. See `.clasp.json` format below.

### 2. Create `.clasp.json`

This file is excluded from the repo (see `.gitignore`) to protect the script ID. Create it manually in the project root:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "./src"
}
```

Find your script ID in the Apps Script editor under **Project Settings**.

### 3. Set Script Properties

In the Apps Script editor, go to **Project Settings → Script Properties** and add the following:

| Key | Required | Description |
|-----|----------|-------------|
| `MASTER_SPREADSHEET_ID` | Only for `copyValidationToAllMonthsAllCampuses` utility | The Google Sheets ID of the master spreadsheet. All other functions use `getActiveSpreadsheet()` since the script is bound. |
| `CONSOLIDATE_BATCH_SIZE` | Optional | Number of campuses per batch. Defaults to `15` if not set. Set via `setConsolidationBatchSize(n)` in the script editor. |

### 4. Set up the control sheet

Populate `CampusBMCSheetInfo` in the master spreadsheet with one row per campus (columns A–D). Column E is filled automatically when you run **Create Campus Spreadsheets**.

### 5. Set up triggers (for automated overnight processing)

Triggers are created and removed automatically by the script when you use the **Auto ES / MS / HS** or **Auto All Levels** menu options. No manual trigger setup is needed for normal use.

If you want to set up a manual time-based trigger for overnight runs, create one in the Apps Script editor pointing to `autoConsolidateAllLevels`.

## Menu reference

The script adds a **🚩 BMC** menu to the master spreadsheet:

| Menu item | What it does |
|-----------|-------------|
| **Start ES / MS / HS** | Resets the cursor for that level, clears existing data rows for that level in the master's monthly tabs, and processes the first batch |
| **Next Batch ES / MS / HS** | Processes the next batch for that level, resuming from where it left off |
| **Auto ES / MS / HS** | Starts automated trigger-based processing for one level |
| **Auto All Levels** | Runs ES → MS → HS automatically in sequence |
| **Stop Auto Processing** | Cancels automation and removes all triggers |
| **Show Status** | Displays per-level progress and the current batch size |
| **Create Campus Spreadsheets** | Creates one spreadsheet per campus from the control sheet |

## Consolidation behavior

- **Start** clears all existing rows for that level across all monthly tabs in the master, then appends fresh data from the first batch.
- **Next Batch** appends to whatever is already in the master — it does not re-clear.
- Rows are validated before writing: invalid campus names (column D) or service type values (column F) are skipped and reported.
- Non-breaking spaces and smart quotes in string values are normalized automatically.

## Batch size

Default is 15 campuses per batch. To change it:

```javascript
// Run this in the Apps Script editor
setConsolidationBatchSize(10);
```

The value persists in Script Properties across runs.

## Utility scripts (`src/utils/`)

These are one-time maintenance functions. Run them from the Apps Script editor as needed:

| File | Function | Purpose |
|------|----------|---------|
| `copyValidationToAllMonthsAllCampuses.js` | `copyValidationToAllMonthsAllCampuses` | Fills campus name column and applies service type validation in all campus spreadsheets. Requires `MASTER_SPREADSHEET_ID` Script Property. |
| `prependYearToCampusSpreadsheetNames.js` | `prependYearToCampusSpreadsheetNames` | Prepends a year-range prefix to campus spreadsheet names. Update the prefix string before running. |
| `removeDataValidation.js` | `modifyDataValidationToAllowCustomEntries` | Switches the secondary disability column from strict dropdown to "show warning" mode, allowing free-text entries. |
| `removeDropdownsAndClearCells.js` | `removeDropdowns` | Removes dropdown validation from the secondary disability column in all campus spreadsheets. |
| `sep25Updates.js` | `updateCampusFilesWithSep25Updates` | Updates column headers, borders, and validation lists in all campus spreadsheets. Update header values and date ranges before running. |
| `updateColNInMonthSheets.js` | `updateColNInMonthSheets` | Updates the cumulative-minutes header cell in each monthly tab. Update the date-range strings before running. |

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| Validation error on write | Confirm headers/validations are in rows 1–2 of the master's monthly tabs and data starts on row 3 |
| Timeouts | Reduce batch size with `setConsolidationBatchSize(n)` and run more batches |
| Missing data after consolidation | Confirm campus monthly tabs exist and contain data starting at the correct row |
| Permission errors on campus spreadsheets | Confirm the Apps Script account has access to the spreadsheet IDs in column E |
| "No files to push" with clasp | Confirm `rootDir` in `.clasp.json` is set to `"./src"` |

## Attribution

Originally created by [Alvaro Gomez](https://github.com/alvarogomez) for Northside Independent School District (NISD). Generalized and released for public use.

Licensed under the [MIT License](LICENSE).
