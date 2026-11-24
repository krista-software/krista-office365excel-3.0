# Get Table Rows

## Overview

Retrieves all rows of data from an Excel table in a file stored in a shared drive. This operation returns the complete data set from the table, including all column values for each row.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the table | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Table Name | Text | Yes | The name of the table to retrieve rows from | "EmployeeTable" |

### Parameter Details

#### Drive Id
- **Format**: Unique identifier for the shared drive
- **How to obtain**: Use "Get All Shared Drives Information" catalog request
- **Validation**: Must be a valid drive ID accessible by the authenticated user
- **Example**: `b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8`

#### File Id
- **Format**: Unique identifier for the Excel file
- **How to obtain**: Use "Get Shared Excel Files" catalog request
- **Validation**: Must be a valid Excel file ID in the specified drive
- **Example**: `01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR`

#### Table Name
- **Format**: Name of the table as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Table must exist in the Excel file
- **Example**: `Table1`, `EmployeeData`, `SalesRecords`

> **📝 Note**: Table names are different from sheet names. Use the actual table name, not the sheet name.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Table Rows | List&lt;Composite&gt; | Array of rows, where each row is an array of cell values |

### Output Structure

The output is a two-dimensional array where:
- Each element in the outer array represents a row
- Each element in the inner array represents a cell value in that row
- Header row is typically included as the first row

### Output Example

```json
{
  "Table Rows": [
    ["Employee ID", "Name", "Department", "Salary"],
    ["EMP001", "John Doe", "Engineering", "75000"],
    ["EMP002", "Jane Smith", "Marketing", "65000"],
    ["EMP003", "Bob Johnson", "Sales", "70000"]
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and user permissions |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Table not found | "Table '{tableName}' not found in the file" | Verify table name (case-sensitive) |
| Table is empty | "Table contains no data" | Ensure table has at least one data row |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters (Drive Id, File Id, Table Name)
- Empty or whitespace-only parameters

**Error Message**: `"Drive ID is required"`, `"File ID is required"`, `"Table Name is required"`

**Resolution**:
1. Verify all required parameters are provided
2. Ensure no parameters are empty or null
3. Check parameter values are valid

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Table doesn't exist in the specified file
- File not accessible or doesn't exist
- Drive not accessible
- Table has no data rows
- Insufficient permissions to read file

**Error Message**: `"Table '{tableName}' not found"`, `"Table contains no data"`

**Resolution**:
1. Verify table name is correct and case-sensitive
2. Ensure file exists and is accessible
3. Check table contains data rows
4. Verify user has read permissions to the file

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)
- Large table causing timeout

**Error Message**: `"Failed to retrieve table rows"`, `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited
5. For very large tables, consider using pagination or filtering

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions
- User doesn't have access to the file

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite.All permission
3. Check user has read access to the file

## Usage Examples

### Example 1: Get All Employee Data

**Scenario**: Retrieve all employee records from table

**Table Structure**:
| Employee ID | Name | Department | Salary |
|-------------|------|------------|--------|
| EMP001 | John Doe | Engineering | 75000 |
| EMP002 | Jane Smith | Marketing | 65000 |
| EMP003 | Bob Johnson | Sales | 70000 |

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Table Name: "EmployeeTable"
```

**Output**:
```json
{
  "Table Rows": [
    ["Employee ID", "Name", "Department", "Salary"],
    ["EMP001", "John Doe", "Engineering", "75000"],
    ["EMP002", "Jane Smith", "Marketing", "65000"],
    ["EMP003", "Bob Johnson", "Sales", "70000"]
  ]
}
```

**Result**: Complete table data retrieved including headers

### Example 2: Get Product Catalog

**Scenario**: Retrieve all products from product table

**Input**:
```
Table Name: "ProductTable"
```

**Output**:
```json
{
  "Table Rows": [
    ["Product Code", "Product Name", "Price", "Stock"],
    ["PRD001", "Widget A", "29.99", "150"],
    ["PRD002", "Widget B", "39.99", "200"],
    ["PRD003", "Widget C", "49.99", "100"]
  ]
}
```

**Result**: All product data retrieved

### Example 3: Empty Table

**Scenario**: Retrieve rows from table with no data

**Input**:
```
Table Name: "EmptyTable"
```

**Output**:
```json
{
  "Table Rows": [
    ["Column1", "Column2", "Column3"]
  ]
}
```

**Result**: Only header row returned (table has no data rows)

### Example 4: Large Table

**Scenario**: Retrieve data from table with many rows

**Input**:
```
Table Name: "TransactionTable"
```

**Output**:
```json
{
  "Table Rows": [
    ["Transaction ID", "Date", "Amount", "Status"],
    ["TXN001", "2024-01-15", "1000.00", "Complete"],
    ["TXN002", "2024-01-16", "2500.00", "Complete"],
    ... (many more rows)
  ]
}
```

**Result**: All rows retrieved (may take longer for large tables)

## Business Rules

1. **Includes Headers**: Header row is included as the first row in the result
2. **All Rows**: Returns all rows from the table (no filtering)
3. **All Columns**: Returns all columns from the table
4. **Preserves Order**: Rows are returned in their original order
5. **Empty Cells**: Empty cells are returned as empty strings
6. **Formula Values**: Returns calculated values, not formulas
7. **Data Types**: All values are returned as text strings
8. **No Formatting**: Cell formatting is not included

## Limitations

1. **Large Tables**: Very large tables may cause timeouts (recommended: <10,000 rows)
2. **No Filtering**: Cannot filter rows (returns all rows)
3. **No Column Selection**: Cannot select specific columns (returns all columns)
4. **No Sorting**: Results are not sorted (returned in table order)
5. **No Formatting**: Cell formatting, colors, and styles are not returned
6. **No Formulas**: Cannot retrieve formula text, only calculated values
7. **Text Representation**: All values are returned as text
8. **Shared Drives Only**: This request is for shared drive files only

## Best Practices

### 1. Handle Large Tables
For tables with many rows, implement pagination or filtering on the client side.

### 2. Parse Header Row
Use the first row to identify column positions for data processing.

### 3. Type Conversion
Convert string values to appropriate types (number, date, boolean) as needed.

### 4. Handle Empty Tables
Check if table has data rows before processing (more than just header row).

### 5. Cache Results
If accessing the same table repeatedly, cache results to reduce API calls.

### 6. Validate Table Exists
Verify table exists before attempting to retrieve rows.

## Common Use Cases

### 1. Data Export
```
Scenario: Export table data to external system
Action: Retrieve all table rows
Result: Data available for processing or storage
```

### 2. Report Generation
```
Scenario: Generate report from table data
Action: Get all rows and format for report
Result: Report created with table data
```

### 3. Data Analysis
```
Scenario: Analyze sales data from table
Action: Retrieve all sales records
Result: Data available for analysis and calculations
```

### 4. Data Synchronization
```
Scenario: Sync table data with external database
Action: Retrieve all rows and compare with database
Result: Data synchronized between systems
```

## Related Catalog Requests

- [Add Rows to Table](pages/AddRowsToTableShared.md) - Add rows to table
- [Table Vlookup](pages/TableVlookup.md) - Lookup specific values
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md) - Get specific columns for keys
- [Get Rows for Range](pages/GetRowsForRange.md) - Get data from sheet range
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get File ID

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation and table existence check
- **Service**: Delegates to WorkbookService for table operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/tables/{tableName}/rows`
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **ROWS_RETRIEVED_COUNT**: Number of rows retrieved
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Table not found" Error

**Cause**: Table name doesn't match or table doesn't exist

**Solution**:
1. Verify table name is spelled correctly (case-sensitive)
2. Check table exists in the Excel file
3. Use table name, not sheet name
4. Verify file contains the specified table

### Issue: Only Header Row Returned

**Cause**: Table has no data rows

**Solution**:
1. Verify table contains data rows
2. Check if data was accidentally deleted
3. Ensure table is not filtered in Excel

### Issue: Timeout on Large Table

**Cause**: Table has too many rows

**Solution**:
1. This is expected for very large tables
2. Consider using filtering or pagination strategies
3. Retrieve data in smaller chunks if possible

### Issue: Missing Columns

**Cause**: Table structure changed or columns hidden

**Solution**:
1. Verify table structure in Excel
2. Check if columns are hidden
3. Ensure table definition includes all expected columns

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Add rows to table
- [Table Vlookup](pages/TableVlookup.md) - Lookup values
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md) - Get specific columns

