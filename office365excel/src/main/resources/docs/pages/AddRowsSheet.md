# Add Rows (Sheet)

## Overview

Adds rows of data to a worksheet in an Excel file stored in the user's personal OneDrive. Data can be provided directly as text rows or uploaded from CSV files. This operation appends rows to the end of the existing data in the sheet.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | The ID of the Excel file | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | The name of the worksheet to add rows to | "Sheet1" |
| Rows | List&lt;Text&gt; | No | List of comma-separated row data | ["John,Doe,30", "Jane,Smith,25"] |
| Rows from Csv | List&lt;File&gt; | No | CSV files containing rows to add | [data.csv] |

### Parameter Details

#### File ID
- **Format**: Unique identifier for the Excel file in OneDrive
- **How to obtain**: Use "Get All Excel Files For Name" catalog request
- **Validation**: Must be a valid Excel file ID accessible by the authenticated user
- **Example**: `01C8MOQGI3YRBDXAQERECWE3W4XCF`

#### Sheet Name
- **Format**: Name of the worksheet as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Sheet must exist in the Excel file
- **Special Characters**: Supported, but avoid special characters when possible
- **Example**: `Sheet1`, `Sales Data`, `Q1 Report`

#### Rows
- **Format**: List of comma-separated text strings
- **Optional**: At least one of "Rows" or "Rows from Csv" must be provided
- **Delimiter**: Comma (,) separates column values
- **Escaping**: Use quotes for values containing commas
- **Example**: `["John,Doe,30,Engineer", "Jane,Smith,25,Manager"]`

> **📝 Note**: Each string represents one row. Values are separated by commas.

#### Rows from Csv
- **Format**: CSV files
- **Optional**: At least one of "Rows" or "Rows from Csv" must be provided
- **Multiple Files**: Supports multiple file uploads
- **File Type**: .csv (comma-separated values)
- **Encoding**: UTF-8 recommended
- **Example**: Upload `employees.csv` containing row data

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sheet Data | List&lt;Composite&gt; | The data that was added to the sheet |

### Output Example

```json
{
  "Sheet Data": [
    ["John", "Doe", "30", "Engineer"],
    ["Jane", "Smith", "25", "Manager"]
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| File ID is empty or null | "File ID is required" | Provide a valid File ID |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Both Rows and Rows from Csv are empty | "Either Rows or Rows from Csv must be provided" | Provide at least one data source |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Sheet not found | "Sheet '{sheetName}' not found in the file" | Verify sheet name (case-sensitive) |
| Invalid file format | "Invalid file format. Only CSV files are supported" | Upload CSV files only |
| File read error | "Failed to read file: {fileName}" | Verify file is not corrupted |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters (File ID, Sheet Name)
- Both Rows and Rows from Csv are empty
- Invalid file format (not CSV)
- Empty or whitespace-only parameters

**Error Message**: `"File ID is required"`, `"Either Rows or Rows from Csv must be provided"`

**Resolution**:
1. Verify all required parameters are provided
2. Provide at least one data source (Rows or Rows from Csv)
3. Ensure files are in CSV format
4. Check no parameters are empty or null

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Sheet doesn't exist in the specified file
- File not accessible or doesn't exist
- Insufficient permissions to modify file
- File is locked or in use by another process

**Error Message**: `"Sheet '{sheetName}' not found"`, `"File not found or inaccessible"`

**Resolution**:
1. Verify sheet name is correct and case-sensitive
2. Ensure file exists and is accessible
3. Check user has write permissions to the file
4. Verify file is not locked or checked out

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)
- File size limits exceeded

**Error Message**: `"Failed to add rows to sheet"`, `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited
5. Reduce batch size if file size limits exceeded

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions
- User doesn't have write access to the file
- File is in a location user cannot modify

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite permission
3. Check user has write access to the file
4. Ensure file is not read-only

## Usage Examples

### Example 1: Add Rows Using Text List

**Scenario**: Add employee data directly as text rows

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Employees"
Rows: ["EMP001,John Doe,Engineering,75000", "EMP002,Jane Smith,Marketing,65000"]
Rows from Csv: null
```

**Output**:
```json
{
  "Sheet Data": [
    ["EMP001", "John Doe", "Engineering", "75000"],
    ["EMP002", "Jane Smith", "Marketing", "65000"]
  ]
}
```

**Result**: Two employee records added to the Employees sheet

### Example 2: Add Rows from CSV File

**Scenario**: Upload employee data from a CSV file

**CSV File Content** (employees.csv):
```
EMP003,Bob Johnson,Sales,70000
EMP004,Alice Williams,HR,60000
EMP005,Charlie Brown,IT,80000
```

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Employees"
Rows: null
Rows from Csv: [employees.csv]
```

**Output**:
```json
{
  "Sheet Data": [
    ["EMP003", "Bob Johnson", "Sales", "70000"],
    ["EMP004", "Alice Williams", "HR", "60000"],
    ["EMP005", "Charlie Brown", "IT", "80000"]
  ]
}
```

**Result**: Three employee records from CSV file added to sheet

### Example 3: Combine Text Rows and File Upload

**Scenario**: Add rows from both direct input and file upload

**Input**:
```
Rows: ["2024-01-15,Product A,100,2999.99"]
Rows from Csv: [sales_data.csv]
```

**Result**: Rows from both sources are combined and added to the sheet

### Example 4: Add Rows with Special Characters

**Scenario**: Add data containing commas in values

**Input**:
```
Rows: ["\"Smith, John\",\"New York, NY\",30"]
```

**Result**: Values with commas are properly handled using quotes

## Business Rules

1. **Append Only**: Rows are always appended to the end of existing data
2. **No Overwrite**: Existing data in the sheet is never overwritten
3. **Column Matching**: Number of values in each row should match sheet columns
4. **Data Types**: All values are treated as text initially; Excel may auto-convert
5. **File Combination**: If both Rows and Rows from Csv are provided, both are processed
6. **Order Preserved**: Rows are added in the order provided
7. **Empty Rows**: Empty rows are skipped
8. **Delimiter**: Comma (,) is the delimiter for text rows and CSV files

## Limitations

1. **Batch Size**: Large batches may hit API limits (recommended: <1000 rows per request)
2. **File Size**: Uploaded files should be under 4MB
3. **No Formatting**: Cell formatting is not applied (plain text only)
4. **No Formulas**: Cannot add formulas, only values
5. **No Validation**: Excel data validation rules are not enforced
6. **Personal OneDrive Only**: This request is for personal OneDrive files only
7. **No Insert**: Cannot insert rows at specific positions (append only)
8. **CSV Only**: Only CSV files supported (not TSV or other formats)

## Best Practices

### 1. Validate Data Before Adding
Ensure data format matches the sheet structure before adding rows.

### 2. Use Batch Operations
For large datasets, split into smaller batches to avoid timeouts.

### 3. Handle Errors Gracefully
Implement error handling for file read errors and API failures.

### 4. Verify Sheet Exists
Check sheet exists before attempting to add rows.

### 5. Use CSV for Large Datasets
For large amounts of data, use CSV files instead of text rows for better performance.

### 6. Escape Special Characters
Use quotes to escape commas and special characters in text values.

## Common Use Cases

### 1. Bulk Data Import
```
Scenario: Import employee data from external system
Action: Upload CSV file with employee records
Result: All employee data added to Excel file
```

### 2. Automated Reporting
```
Scenario: Add daily sales data to report
Action: Provide sales data as text rows
Result: Daily sales appended to report sheet
```

### 3. Data Migration
```
Scenario: Migrate data from legacy system to Excel
Action: Export data to CSV and upload to Excel
Result: Legacy data available in Excel format
```

### 4. Form Submissions
```
Scenario: Add form submission data to tracking sheet
Action: Each submission adds a new row
Result: All submissions tracked in Excel
```

## Related Catalog Requests

- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Shared drive version
- [Add Rows](pages/AddRows.md) - Table version
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Find File ID
- [Get Cell Value](pages/GetCellValue.md) - Read cell values
- [Delete Sheet](pages/DeleteSheet.md) - Delete sheets

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation and file processing
- **Service**: Delegates to WorksheetService for sheet operations

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range`
- **Permissions Required**: Files.ReadWrite

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **ROWS_ADDED_COUNT**: Number of rows added
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Sheet not found" Error

**Cause**: Sheet name doesn't match or sheet doesn't exist

**Solution**:
1. Verify sheet name is spelled correctly (case-sensitive)
2. Use "Get Sheets Information" to list available sheets
3. Check for extra spaces in sheet name

### Issue: Data Not Formatted Correctly

**Cause**: Excel auto-formatting or data type conversion

**Solution**:
1. This is expected Excel behavior
2. Pre-format cells in Excel if specific formatting is needed
3. Use text format for cells that should remain as text

### Issue: File Upload Fails

**Cause**: File format or size issues

**Solution**:
1. Verify file is CSV format
2. Check file size is under 4MB
3. Ensure file encoding is UTF-8
4. Verify file is not corrupted

### Issue: Rows Added to Wrong Location

**Cause**: Rows are always appended to the end

**Solution**:
This is expected behavior. Use "Add Row At Index" if you need to insert at a specific position.

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Shared drive version
- [Add Rows](pages/AddRows.md) - Table version
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Find files

