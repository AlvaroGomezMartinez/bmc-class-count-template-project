# BMC Class Count Template Project

## Product Overview

This Google Apps Script is bound to the main BMC Class Count Template spreadsheet and automates the consolidation of monthly BMC (Behavior Management Classroom) data from individual campus spreadsheets into a centralized dashboard for Northside Independent School District.

## Key Features

- **Campus Spreadsheet Creation**: Automatically creates individual campus spreadsheets from a master template, organized by school level (ES/MS/HS)
- **Batch Data Consolidation**: Processes campus data in configurable batches (default 15) to avoid timeouts and API limits
- **Level-Based Processing**: Handles Elementary (ES), Middle School (MS), and High School (HS) data separately with resumable batch processing
- **Automated Processing**: Supports overnight automated consolidation with email notifications and error handling
- **Data Validation**: Maintains spreadsheet validations and handles schema alignment between campus and master sheets

## Target Users

- District special education administrators, coordinators, specialists and secretaries managing BMC class count data
- Campus BMC teachers who input monthly data

**Access Requirements**: All users must be logged into the Northside Independent School District network to access and use the system.

## Data Flow

1. Control sheet (`CampusBMCSheetInfo`) contains campus information and spreadsheet IDs
2. Individual campus spreadsheets collect monthly data (August through April/May projections)
3. Consolidation process aggregates campus data back into master spreadsheet's monthly tabs
4. Data is organized by school level and processed in safe, resumable batches