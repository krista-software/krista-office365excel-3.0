# Add Rows to Table

## Overview

Adds one or more rows to an existing table in a shared Excel file stored in OneDrive or SharePoint. This operation supports both direct row input and CSV file upload for bulk data insertion.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the table | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Table Name | Text | Yes | The name of the table where rows will be added | "Table1" |
| Rows | List&lt;Text&gt; | No* | List of comma-separated row data to add | ["Data1,Data2,Data3", "Data4,Data5,Data6"] |
| Rows from Csv File | List&lt;File&gt; | No* | CSV file(s) containing rows to be added to the table | CSV file with row data |

**\*Note**: Either `Rows` or `Rows from Csv File` must be provided. Both parameters are optional individually, but at least one must contain data.

### Parameter Details

#### Drive Id
- **Format**: Unique identifier for the OneDrive or SharePoint drive
- **How to obtain**: Use "Get All Shared Drives Information" or "Get Shared Excel Files" catalog request
- **Validation**: Must be a valid drive ID accessible by the authenticated user
- **Example**: `b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8`

#### File Id
- **Format**: Unique identifier for the Excel file
- **How to obtain**: Use "Get Shared Excel Files" catalog request
- **Validation**: Must be a valid Excel file ID within the specified drive
- **Example**: `01C8MOQGI3YRBDXAQERECWE3W4XCF`

#### Table Name
- **Format**: Name of the table as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Table must exist in the Excel file
- **Example**: `Table1`, `SalesData`, `EmployeeRecords`

#### Rows
- **Format**: List of comma-separated strings, each representing a row
- **Structure**: Each string should contain values separated by commas matching table columns
- **Column Count**: Must match the number of columns in the table
- **Example**: `["John,Doe,30,Engineer", "Jane,Smith,28,Designer"]`

#### Rows from Csv File
- **Format**: CSV file with proper formatting
- **Structure**: Each line represents a row, values separated by commas
- **Headers**: Should not include header row (only data rows)
- **Multiple Files**: Supports uploading multiple CSV files
- **Example**: File containing:
  ```
  John,Doe,30,Engineer
  Jane,Smith,28,Designer
  ```

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Rows | List&lt;Composite&gt; | The rows that were successfully added to the table |

### Output Structure

The output contains a list of composite objects representing the added rows with their values.

**Example Output**:
```json
{
  "Rows": [
    {
      "Column1": "John",
      "Column2": "Doe",
      "Column3": "30",
      "Column4": "Engineer"
    },
    {
      "Column1": "Jane",
      "Column2": "Smith",
      "Column3": "28",
      "Column4": "Designer"
    }
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Both Rows and Rows from Csv File are empty | "Row data not provided" | Provide data in either Rows or Rows from Csv File parameter |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and user permissions |
| File not found | "File not found in the specified drive" | Verify File ID and ensure file exists |
| Table not found | "Table '{tableName}' not found in the file" | Verify table name (case-sensitive) |
| Column count mismatch | "Row has {actual} columns but table expects {expected}" | Ensure each row has the correct number of columns |
| Invalid CSV format | "Invalid CSV file format" | Verify CSV file is properly formatted |
| File size exceeds limit | "File size exceeds maximum allowed size" | Reduce file size or split into multiple requests |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters (Drive Id, File Id, Table Name)
- No row data provided (both Rows and Rows from Csv File are empty)
- Invalid parameter formats

**Error Message**: `"Row data not provided"` or `"[Parameter Name] is required"`

**Resolution**:
1. Verify all required parameters are provided
2. Ensure at least one data source (Rows or CSV file) contains data
3. Check parameter formats match expected types

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Table doesn't exist in the specified file
- Column count mismatch between input data and table structure
- Drive or file not accessible
- Insufficient permissions

**Error Message**: `"Table '{tableName}' not found in the file"` or `"Row has {actual} columns but table expects {expected}"`

**Resolution**:
1. Verify table name is correct and case-sensitive
2. Ensure row data matches table column count
3. Check user has write permissions to the file
4. Confirm file and drive are accessible

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to add rows to table"` or `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions
- User doesn't have access to the file
- File is read-only

**Error Message**: `"Unauthorized access"` or `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite.All permission
3. Check user has edit access to the file
4. Ensure file is not read-only

## Usage Examples

### Example 1: Add Rows Using Direct Input

**Scenario**: Add employee records to a table

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Rows: ["John,Doe,Engineering,50000", "Jane,Smith,Marketing,55000"]
```

**Output**:
```json
{
  "Rows": [
    {
      "FirstName": "John",
      "LastName": "Doe",
      "Department": "Engineering",
      "Salary": "50000"
    },
    {
      "FirstName": "Jane",
      "LastName": "Smith",
      "Department": "Marketing",
      "Salary": "55000"
    }
  ]
}
```

**Result**: Two employee records successfully added to EmployeeTable

### Example 2: Add Rows Using CSV File

**Scenario**: Bulk import sales data from CSV file

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "SalesData"
Rows from Csv File: [sales_data.csv]
```

**CSV File Content** (sales_data.csv):
```
2024-01-15,Product A,100,5000
2024-01-16,Product B,150,7500
2024-01-17,Product C,200,10000
```

**Output**:
```json
{
  "Rows": [
    {"Date": "2024-01-15", "Product": "Product A", "Quantity": "100", "Revenue": "5000"},
    {"Date": "2024-01-16", "Product": "Product B", "Quantity": "150", "Revenue": "7500"},
    {"Date": "2024-01-17", "Product": "Product C", "Quantity": "200", "Revenue": "10000"}
  ]
}
```

**Result**: Three sales records successfully imported from CSV file

### Example 3: Retry Scenario

**Scenario**: Operation fails due to rate limiting, automatic retry succeeds

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "DataTable"
Rows: ["Value1,Value2,Value3"]
```

**Process**:
1. Initial request encounters rate limiting (429 error)
2. Extension automatically waits and retries
3. Retry succeeds and rows are added

**Output**:
```json
{
  "Rows": [
    {"Column1": "Value1", "Column2": "Value2", "Column3": "Value3"}
  ]
}
```

**Result**: Row successfully added after automatic retry

## Business Rules

1. **Column Matching**: The number of values in each row must exactly match the number of columns in the target table
2. **Data Types**: Excel will attempt to auto-detect data types; explicit formatting may be needed for dates, numbers, etc.
3. **Table Expansion**: The table automatically expands to accommodate new rows
4. **Append Operation**: Rows are always appended to the end of the table
5. **Atomic Operation**: All rows are added in a single operation; partial failures are not supported
6. **CSV Format**: CSV files must use comma as delimiter; other delimiters are not supported
7. **File Encoding**: CSV files should use UTF-8 encoding for proper character support

## Limitations

1. **Maximum Rows**: Limited by Microsoft Graph API constraints (typically 5,000 rows per request)
2. **File Size**: CSV files are limited to 100 MB per file
3. **Concurrent Operations**: Multiple simultaneous writes to the same table may cause conflicts
4. **Rich Text**: Rich text formatting is not supported; only plain text values
5. **Formulas**: Input data is treated as values, not formulas
6. **Special Characters**: Commas within cell values must be properly escaped in CSV files
7. **Rate Limiting**: Subject to Microsoft Graph API rate limits (automatic retry handles this)

## Best Practices

### 1. Use CSV Files for Large Datasets
For adding more than 100 rows, use CSV file upload instead of direct row input for better performance and reliability.

### 2. Validate Data Before Upload
Ensure all rows have the correct number of columns and proper formatting before submitting the request.

### 3. Handle Errors Gracefully
Implement error handling in your workflow to manage validation failures and retry logic.

### 4. Monitor Rate Limits
Be aware of Microsoft Graph API rate limits when performing bulk operations. The extension handles retries automatically.

### 5. Use Descriptive Table Names
Use clear, descriptive table names to avoid confusion and make maintenance easier.

### 6. Test with Small Datasets First
When implementing new workflows, test with a small number of rows before processing large datasets.

## Common Use Cases

### 1. Employee Onboarding
```
Scenario: Add new employee records to HR tracking table
Action: Use direct row input with employee details
Result: New employees appear in the tracking table immediately
```

### 2. Sales Data Import
```
Scenario: Import daily sales data from external system
Action: Export sales data to CSV, upload using Rows from Csv File
Result: Sales data synchronized with Excel tracking sheet
```

### 3. Inventory Updates
```
Scenario: Bulk update inventory levels from warehouse system
Action: Generate CSV with current inventory counts, upload to inventory table
Result: Inventory table reflects current stock levels
```

## Related Catalog Requests

- [Add Rows to Table (Async)](pages/AddRowsToTableAsync.md) - Asynchronous version for large datasets
- [Get Add Rows To Table Result](pages/GetAddRowsToTableResult.md) - Retrieve results of async operation
- [Get Table Rows](pages/GetTableRows.md) - Retrieve existing rows from a table
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Find File ID and Drive ID
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Find Drive ID

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation performed before API call
- **Service**: Delegates to SharedTableResponseService for table operations

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Table not found" Error

**Cause**: Table name doesn't match or table doesn't exist

**Solution**:
1. Verify table name is spelled correctly (case-sensitive)
2. Use "Get Sheets Information" to list available tables
3. Ensure you're using the table name, not the sheet name

### Issue: Column Count Mismatch

**Cause**: Row data has different number of columns than table

**Solution**:
1. Count the number of columns in your table
2. Ensure each row has exactly that many comma-separated values
3. Check for extra commas or missing values

### Issue: CSV File Not Processing

**Cause**: Invalid CSV format or encoding

**Solution**:
1. Ensure CSV uses comma as delimiter
2. Save CSV with UTF-8 encoding
3. Remove any header rows (only data rows should be included)
4. Verify file is not corrupted

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add rows to a sheet (not a table)
- [Add Rows](pages/AddRowsTable.md) - Table Area version for personal OneDrive files

