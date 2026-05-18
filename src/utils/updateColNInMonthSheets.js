/**
 * Updates the cumulative-minutes header cell in each monthly sheet of every
 * campus spreadsheet listed in CampusBMCSheetInfo column E.
 *
 * Each month maps to a specific cell (N2 for regular months, N3 for the
 * projections tab) and a date-range label string. Update the values in
 * monthUpdates to match the current school year's date ranges before running.
 *
 * @returns {void}
 */
function updateColNInMonthSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName('CampusBMCSheetInfo');
    const sheetIds = infoSheet.getRange('E2:E' + infoSheet.getLastRow()).getValues().flat().filter(Boolean);

    // Map of month sheet names to their new N2/N3 values and target cell
    const monthUpdates = {
        'AUGUST':        { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - AUGUST 29, 2025' },
        'SEPTEMBER':     { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - SEPTEMBER 30, 2025' },
        'OCTOBER':       { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - OCTOBER 31, 2025' },
        'NOVEMBER':      { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - NOVEMBER 21, 2025' },
        'DECEMBER':      { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - DECEMBER 19, 2025' },
        'JANUARY':       { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - JANUARY 30, 2026' },
        'FEBRUARY':      { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - FEBRUARY 27, 2026' },
        'MARCH':         { cell: 'N2', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - MARCH 31, 2026' },
        'APRIL/ MAY PROJECTIONS': { cell: 'N3', value: 'CUMULATIVE MINUTES: START DATE AUGUST 11, 2025 - APRIL 30, 2026' }
    };

    sheetIds.forEach(id => {
        const extSS = SpreadsheetApp.openById(id);
        Object.entries(monthUpdates).forEach(([sheetName, {cell, value}]) => {
            const sheet = extSS.getSheetByName(sheetName);
            if (sheet) {
                sheet.getRange(cell).setValue(value);
            }
        });
    });
}