# BMC Class Count Project

## Overview
This Google Apps Script automates the consolidation of monthly data from all campus BMCs into a main (dashboard) spreadsheet. Data is collected in batches by school level (ES/MS/HS).

- Creates individual campus spreadsheets using data from the control sheet called <em>CampusBMCSheetInfo</em>.
- Consolidates each campus’ monthly sheets back into the main spreadsheet's sheet called <em>Information Regarding Class Counts</em>. The function is run by level in safe, resumable batches.
- Preserves headers/validations in the monthly sheets; data writes start on row 3 of each monthly sheet.
- Default batch size is 15 campuses per run.

## Sheets and columns
- Control sheet: `CampusBMCSheetInfo` found in the spreadsheet called <em>BMC Class Count Template</em>
  - Column A: Emails (editor/BMC teacher to share their campus sheet with)
  - Column B: Campus
  - Column C: Level (ES, MS, HS)
  - Column D: Main/Level Folder ID (GDrive folder where the campus spreadsheets are stored)
  - Column E: Campus Spreadsheet ID (individual campus spreadsheets; filled by script when a spreadsheet is created)
- Monthly tabs (sheet) in each campus spreadsheet:
  - `AUGUST`, `SEPTEMBER`, `OCTOBER`, `NOVEMBER`, `DECEMBER`, `JANUARY`, `FEBRUARY`, `MARCH`, `APRIL/ MAY PROJECTIONS`
  - For the months of <em>August - March</em>, data begins on row 3 (rows 1–2 reserved for headers/metadata/validations).
  - For the month of <em>April/ May Projections</em> data begins on row 4 (rows 1-3 reserved for headers/metadata/validations).

## Menu actions (in the main spreadsheet called <em>BMC Class Count Template</em>)
- 🚩 BMC
  - Get Campus Data
    - Start ES / MS / HS: Resets cursor for that level, clears existing rows 3+ across monthly tabs for that level, then processes the first batch.
    - Next Batch ES / MS / HS: Processes the next batch for that level (resumes where it left off).
    - Show Status: Displays per-level progress and the active batch size.
  - Create Campus Spreadsheets: Makes one copy per campus, moves the spreadsheet into the level (ES, MS, HS) folder, gets the spreadsheet ID and adds it to column E in the <em>CampusBMCSheetInfo</em> sheet.

## Consolidation behavior
- Batching: Processes campuses in groups (default 15). Use `Next Batch...` repeatedly until status shows DONE.
- Overwrite semantics: When you click Start for a level, rows 3+ for that level are cleared in the main spreadsheet's months tabs; subsequent batches append fresh data aligned to existing columns.
- Schema safety: The script doesn’t modify headers or insert columns; it trims/pads incoming rows to the master’s current column count to avoid breaking validations.

## Batch size
- Default: 15. You can change it via the function `setConsolidationBatchSize(n)` in the Script Editor (n ≥ 1). The value persists via Script Properties.

## How to use
1) Ensure `CampusBMCSheetInfo` has valid values (especially Level in C and Spreadsheet ID in E when consolidating).
2) In the main spreadsheet menu, choose:
   - Start ES (or MS/HS) → then Next Batch until DONE.
   - Use `Show Status` any time to see progress.
3) To regenerate a level, click `Start...` for that level again; it overwrites rows 3+ for that level and re-appends fresh data.

## Testing tips
- Use a copy of the master to test safely.
- Set a small batch size (e.g., `setConsolidationBatchSize(2)`) and verify behavior with a few campus spreadsheets.
- Edge cases covered: missing spreadsheet ID, missing month sheet, blank rows, invalid IDs/permissions.

## Troubleshooting
- Validation error on write: Ensure your month tabs have headers/validations in rows 1–2 and that data begins on row 3. The script aligns rows to the main spreadsheet’s column count.
- Timeouts: Reduce batch size via `setConsolidationBatchSize(n)` and run more batches.
- Missing data: Confirm campus monthly tabs exist and contain data from row 3 down.
- Permission errors: Make sure the Apps Script’s account can open campus spreadsheet IDs in column E.

## Notes
- The script uses LockService and Script Properties to avoid concurrent runs and to track progress.

## Google Apps Script Development
[Alvaro Gomez](alvaro.gomez@nisd.net), Academic Technology Coach<br>
Department of Academic Technology<br>
Northside Independent School District<br>
