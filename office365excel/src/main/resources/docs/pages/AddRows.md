# Add Rows

## Overview

Adds rows of data to an Excel table in a file stored in the user's personal OneDrive. Data can be provided directly as text rows or uploaded from CSV files. This operation appends rows to the end of the table.

## Request Details

- **Area**: Table
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | The ID of the Excel file containing the table | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Table Name | Text | Yes | The name of the table to add rows to | "Table1" |
| Rows | List&lt;Text&gt; | No | List of comma-separated row data | ["John,Doe,30", "Jane,Smith,25"] |
| Rows From Csv | List&lt;File&gt; | No | CSV files containing rows to add | [data.csv] |

### Parameter Details

#### File ID
- **Format**: Unique identifier for the Excel file in OneDrive
- **How to obtain**: Use "Get All Excel Files For Name" catalog request
- **Validation**: Must be a valid Excel file ID accessible by the authenticated user
- **Example**: `01C8MOQGI3YRBDXAQERECWE3W4XCF`

#### Table Name
- **Format**: Name of the table as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Table must exist in the Excel file
- **Example**: `Table1`, `EmployeeData`, `SalesRecords`

> **📝 Note**: Table names are different from sheet names. Use the actual table name, not the sheet name.

#### Rows
- **Format**: List of comma-separated text strings
- **Optional**: At least one of "Rows" or "Rows From Csv" must be provided
- **Delimiter**: Comma (,) separates column values
- **Escaping**: Use quotes for values containing commas
- **Column Matching**: Number of values must match table columns
- **Example**: `["EMP001,John Doe,Engineering,75000", "EMP002,Jane Smith,Marketing,65000"]`

#### Rows From Csv
- **Format**: CSV files
- **Optional**: At least one of "Rows" or "Rows From Csv" must be provided
- **Multiple Files**: Supports multiple file uploads
- **File Type**: .csv (comma-separated values)
- **Encoding**: UTF-8 recommended
- **Header Row**: Should not include header row (data only)
- **Example**: Upload `employees.csv` containing row data

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Rows | List&lt;Composite&gt; | The rows that were successfully added to the table |

### Output Example

```json
{
  "Rows": [
    ["EMP001", "John Doe", "Engineering", "75000"],
    ["EMP002", "Jane Smith", "Marketing", "65000"]
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| File ID is empty or null | "File ID is required" | Provide a valid File ID |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Both Rows and Rows From Csv are empty | "Row data not provided" | Provide at least one data source |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Table not found | "Table '{tableName}' not found in the file" | Verify table name (case-sensitive) |
| Column count mismatch | "Row has {actual} columns but table has {expected} columns" | Ensure row data matches table structure |
| Invalid file format | "Invalid file format. Only CSV files are supported" | Upload CSV files only |
| File read error | "Failed to read file: {fileName}" | Verify file is not corrupted |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters (File ID, Table Name)
- Both Rows and Rows From Csv are empty
- Invalid file format (not CSV)
- Empty or whitespace-only parameters

**Error Message**: `"File ID is required"`, `"Row data not provided"`

**Resolution**:
1. Verify all required parameters are provided
2. Provide at least one data source (Rows or Rows From Csv)
3. Ensure files are in CSV format
4. Check no parameters are empty or null

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Table doesn't exist in the specified file
- Column count doesn't match table structure
- File not accessible or doesn't exist
- Insufficient permissions to modify file
- File is locked or in use

**Error Message**: `"Table '{tableName}' not found"`, `"Row has {actual} columns but table has {expected} columns"`

**Resolution**:
1. Verify table name is correct and case-sensitive
2. Ensure row data has correct number of columns
3. Check file exists and is accessible
4. Verify user has write permissions to the file
5. Ensure file is not locked

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)
- File size limits exceeded

**Error Message**: `"Failed to add rows to table"`, `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited
5. Reduce batch size if limits exceeded

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions
- User doesn't have write access to the file
- File is read-only

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite permission
3. Check user has write access to the file
4. Ensure file is not read-only

## Usage Examples

### Example 1: Add Rows Using Text List

**Scenario**: Add employee data directly as text rows

**Table Structure**: EmployeeTable with columns: ID, Name, Department, Salary

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Rows: ["EMP001,John Doe,Engineering,75000", "EMP002,Jane Smith,Marketing,65000"]
Rows From Csv: null
```

**Output**:
```json
{
  "Rows": [
    ["EMP001", "John Doe", "Engineering", "75000"],
    ["EMP002", "Jane Smith", "Marketing", "65000"]
  ]
}
```

**Result**: Two employee records added to the table

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
Table Name: "EmployeeTable"
Rows: null
Rows From Csv: [employees.csv]
```

**Output**:
```json
{
  "Rows": [
    ["EMP003", "Bob Johnson", "Sales", "70000"],
    ["EMP004", "Alice Williams", "HR", "60000"],
    ["EMP005", "Charlie Brown", "IT", "80000"]
  ]
}
```

**Result**: Three employee records from CSV file added to table

### Example 3: Combine Text Rows and CSV File

**Scenario**: Add rows from both direct input and file upload

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "SalesTable"
Rows: ["2024-01-15,Product A,100,2999.99"]
Rows From Csv: [sales_data.csv]
```

**Result**: Rows from both sources are combined and added to the table

### Example 4: Column Count Mismatch Error

**Scenario**: Attempt to add row with wrong number of columns

**Table Structure**: 4 columns (ID, Name, Department, Salary)

**Input**:
```
Rows: ["EMP001,John Doe,Engineering"]
```

**Error**: LOGIC_ERROR - "Row has 3 columns but table has 4 columns"

**Result**: Operation fails due to column mismatch

## Business Rules

1. **Append Only**: Rows are always appended to the end of the table
2. **No Overwrite**: Existing data in the table is never overwritten
3. **Column Matching**: Number of values in each row must match table columns
4. **Data Types**: All values are treated as text initially; Excel may auto-convert
5. **File Combination**: If both Rows and Rows From Csv are provided, both are processed
6. **Order Preserved**: Rows are added in the order provided
7. **Empty Rows**: Empty rows are skipped
8. **Table Auto-Expand**: Table automatically expands to accommodate new rows

## Limitations

1. **Batch Size**: Large batches may hit API limits (recommended: <1000 rows per request)
2. **File Size**: Uploaded CSV files should be under 4MB
3. **No Formatting**: Cell formatting is not applied (plain text only)
4. **No Formulas**: Cannot add formulas, only values
5. **No Validation**: Excel data validation rules are not enforced during add
6. **Personal OneDrive Only**: This request is for personal OneDrive files only
7. **No Insert**: Cannot insert rows at specific positions (append only)
8. **CSV Only**: Only CSV files supported (not TSV or other formats)

## Best Practices

### 1. Validate Data Before Adding
Ensure data format matches the table structure before adding rows.

### 2. Match Column Count
Verify the number of values in each row matches the table's column count.

### 3. Use Batch Operations
For large datasets, split into smaller batches to avoid timeouts.

### 4. Handle Errors Gracefully
Implement error handling for file read errors and API failures.

### 5. Verify Table Exists
Check table exists before attempting to add rows.

### 6. Use CSV for Large Datasets
For large amounts of data, use CSV files instead of text rows for better performance.

## Common Use Cases

### 1. Bulk Data Import
```
Scenario: Import employee data from external system
Action: Upload CSV file with employee records
Result: All employee data added to Excel table
```

### 2. Automated Reporting
```
Scenario: Add daily sales data to table
Action: Provide sales data as text rows
Result: Daily sales appended to table
```

### 3. Data Migration
```
Scenario: Migrate data from legacy system to Excel
Action: Export data to CSV and upload to Excel table
Result: Legacy data available in Excel format
```

### 4. Form Submissions
```
Scenario: Add form submission data to tracking table
Action: Each submission adds a new row to table
Result: All submissions tracked in Excel table
```

## Related Catalog Requests

- [Add Rows Async](pages/AddRowsAsync.md) - Asynchronous version for large datasets
- [Get Rows](pages/GetRows.md) - Retrieve rows from table
- [Table Vlookup](pages/TableVlookup.md) - Search for data in table
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Find File ID
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Shared drive version

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation, column count verification, file processing
- **Service**: Delegates to WorkbookService for table operations

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/tables/{tableName}/rows`
- **Permissions Required**: Files.ReadWrite

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **ROWS_ADDED_COUNT**: Number of rows added
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Table not found" Error

**Cause**: Table name doesn't match or table doesn't exist

**Solution**:
1. Verify table name is spelled correctly (case-sensitive)
2. Check table exists in the Excel file
3. Use table name, not sheet name
4. Verify file contains the specified table

### Issue: "Column count mismatch" Error

**Cause**: Row data has different number of columns than table

**Solution**:
1. Count the number of columns in your table
2. Ensure each row has exactly that many values
3. Check for missing or extra commas in row data
4. Verify CSV file structure matches table

### Issue: File Upload Fails

**Cause**: File format or size issues

**Solution**:
1. Verify file is CSV format
2. Check file size is under 4MB
3. Ensure file encoding is UTF-8
4. Verify file is not corrupted

### Issue: Data Not Formatted Correctly

**Cause**: Excel auto-formatting or data type conversion

**Solution**:
1. This is expected Excel behavior
2. Pre-format table columns in Excel if specific formatting is needed
3. Use text format for columns that should remain as text

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Add Rows Async](pages/AddRowsAsync.md) - Asynchronous version
- [Get Rows](pages/GetRows.md) - Retrieve table rows
- [Table Vlookup](pages/TableVlookup.md) - Search table data

