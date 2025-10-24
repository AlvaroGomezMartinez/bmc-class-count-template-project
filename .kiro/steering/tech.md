# Technology Stack & Development Guidelines

## Platform & Runtime

- **Google Apps Script**: V8 runtime environment (bound to main spreadsheet)
- **Google Workspace APIs**: Sheets API, Drive API, Gmail API
- **Network Requirements**: School district intranet access required
- **Authentication**: Users must be logged into district network
- **Timezone**: America/Chicago
- **Exception Logging**: Google Cloud Stackdriver

## Development Tools

- **Script Editor**: Web-based IDE for Google Apps Script (accessed via bound spreadsheet)
- **clasp**: Google Apps Script CLI for local development and deployment
- **Spreadsheet Integration**: Script is bound to the main BMC Class Count Template spreadsheet

## Architecture Patterns

### Batch Processing Pattern
- Use configurable batch sizes (default 15) to prevent API timeouts
- Implement cursor-based resumable processing with Script Properties
- Always use LockService to prevent concurrent executions

### Error Handling & Resilience
- Wrap main functions in try-catch with proper lock management
- Implement exponential backoff for API rate limiting
- Use validation before data writes to prevent corruption
- Log detailed error information for debugging

### Data Validation Strategy
- Clear validation rules before writing data, restore after
- Align row data to destination column count to preserve schema
- Filter invalid data before batch writes
- Use predefined lists for campus names and service types

## Code Conventions

### Function Naming
- Public menu functions: `functionName()` (camelCase)
- Private helper functions: `functionName_()` (underscore suffix)
- Level-specific wrappers: `functionNameES()`, `functionNameMS()`, `functionNameHS()`

### Property Keys
- Use descriptive UPPERCASE keys: `CONSOLIDATE_BATCH_SIZE`, `CONS_LEVEL_IDX_ES`
- Include level suffix for level-specific properties: `_ES`, `_MS`, `_HS`

### Logging Standards
- Log function start/end with timestamps
- Include batch progress information
- Log validation errors with specific row/column details
- Use structured logging for automated processing

## Common Commands

### Development
```javascript
// Set batch size for testing
setConsolidationBatchSize(2)

// Check processing status
showConsolidationStatus()

// Manual batch processing
consolidateLevelStartES()
consolidateLevelNextBatchES()
```

### Deployment
```bash
# Push code to Google Apps Script
clasp push

# View logs
clasp logs
```

## API Limits & Best Practices

- **Batch Size**: Keep under 20 campuses per batch to avoid timeouts
- **Rate Limiting**: Add delays between API calls (1-2 seconds)
- **Retry Logic**: Implement 3-attempt retry with exponential backoff
- **Lock Management**: Always use try-finally to release locks
- **Validation**: Clear/restore data validation rules around writes