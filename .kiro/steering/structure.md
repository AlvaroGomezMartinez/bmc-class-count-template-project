# Project Structure & Organization

## File Organization

**Note**: This is a bound Google Apps Script project attached to the main BMC Class Count Template spreadsheet.

```
/
├── Code.js                  # Main consolidation and menu functions
├── appsscript.json          # Project configuration and runtime settings
├── .clasp.json              # clasp CLI configuration for deployment
├── .claspignore             # Files to exclude from deployment
├── README.md                # Comprehensive project documentation
├── LICENSE                  # Project license
├── utils/                   # Utility functions for maintenance tasks
│   ├── copyValidationToAllMonthsAllCampuses.js
│   ├── prependYearToCampusSpreadsheetNames.js
│   ├── removeDataValidation.js
│   ├── removeDropdownsAndClearCells.js
│   ├── sep25Updates.js
│   └── updateColNInMonthSheets.js
└── tests/                   # Test functions and mock data
    └── testCreateCampusSpreadsheets.js
```

## Core Components

### Main Script (Code.js)

- **Menu System**: `onOpen()` creates custom UI menu with consolidation options
- **Campus Creation**: `createCampusSpreadsheets()` generates individual campus files
- **Batch Processing**: Level-specific consolidation functions (ES/MS/HS)
- **Automation**: Time-based triggers for overnight processing
- **Email Notifications**: Completion and error reporting

### Utility Scripts (utils/)

- **Maintenance Functions**: One-time operations for data updates
- **Validation Management**: Campus name validation and data validation updates
- **Batch Processing**: Similar patterns to main script for consistency

### Test Scripts (tests/)

- **Mock Data Testing**: Simulate operations without affecting live data
- **Logic Validation**: Test business rules and error handling

## Data Structure

### Control Sheet: `CampusBMCSheetInfo`

- **Column A**: Editor emails for sharing
- **Column B**: Campus names
- **Column C**: School level (ES/MS/HS)
- **Column D**: Google Drive folder IDs
- **Column E**: Campus spreadsheet IDs (populated by script)

### Monthly Sheets

- **Regular Months**: `AUGUST` through `MARCH` (data starts row 3)
- **Projections**: `APRIL/ MAY PROJECTIONS` (data starts row 4)
- **Headers**: Rows 1-2 contain metadata and validation rules
- **Data Columns**: Standardized across all campus and master sheets

## Naming Conventions

### Files

- Main functionality in root-level `Code.js`
- Utilities prefixed by purpose: `copyValidation...`, `remove...`, `update...`
- Tests prefixed with `test`

### Functions

- Public functions: camelCase without suffix
- Private helpers: camelCase with underscore suffix `_`
- Level wrappers: function name + level (`ES`, `MS`, `HS`)

### Properties & Keys

- Script Properties use UPPERCASE with underscores
- Level-specific keys include level suffix
- Cursor keys follow pattern: `CONS_LEVEL_IDX_[LEVEL]`

## Development Workflow

1. **Local Development**: Use clasp for version control and deployment
2. **Testing**: Use mock data functions before live operations
3. **Batch Operations**: Always implement cursor-based resumable processing
4. **Error Handling**: Comprehensive logging and user feedback
5. **Documentation**: Maintain README.md with usage instructions
