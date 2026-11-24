# Get Column Values For Keys

## Overview

Retrieves specific column values from an Excel table in a shared drive file based on key row values. This operation performs a lookup to find rows matching the provided keys and returns values from specified columns.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the table | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Table Name | Text | Yes | The name of the table containing the data | "EmployeeTable" |
| Key Rows | List&lt;Text&gt; | Yes | List of key values to search for in the first column | ["EMP001", "EMP002"] |
| Lookup Column | List&lt;Text&gt; | Yes | List of column names to retrieve values from | ["Name", "Department"] |
| Get All Record | Boolean | Yes | If true, returns all records; otherwise, returns only matching records | false |

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

#### Key Rows
- **Format**: List of text values to search for
- **Search Column**: Searches in the first column of the table
- **Case Sensitivity**: Case-insensitive matching
- **Multiple Keys**: Can provide multiple keys to find multiple rows
- **Example**: `["EMP001", "EMP002", "EMP003"]`

#### Lookup Column
- **Format**: List of column names to retrieve
- **Case Sensitivity**: Case-sensitive
- **Multiple Columns**: Can retrieve values from multiple columns
- **Validation**: Column names must exist in the table
- **Example**: `["Name", "Department", "Salary"]`

#### Get All Record
- **Format**: Boolean (true/false)
- **True**: Returns all rows from the table (ignores Key Rows)
- **False**: Returns only rows matching the Key Rows
- **Default Behavior**: Use `false` for filtered results
- **Example**: `false`, `true`

> **📝 Note**: When Get All Record = true, the Key Rows parameter is ignored and all table rows are returned.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Result | List&lt;Composite&gt; | Array of rows containing the requested column values |

### Output Structure

Each row in the result contains values from the specified lookup columns.

### Output Example

**When Get All Record = false**:
```json
{
  "Result": [
    ["John Doe", "Engineering"],
    ["Jane Smith", "Marketing"]
  ]
}
```

**When Get All Record = true**:
```json
{
  "Result": [
    ["John Doe", "Engineering"],
    ["Jane Smith", "Marketing"],
    ["Bob Johnson", "Sales"],
    ["Alice Williams", "HR"]
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Key Rows is empty or null | "Key Rows is required" | Provide at least one key value |
| Lookup Column is empty or null | "Lookup Column is required" | Provide at least one column name |
| Get All Record is null | "Get All Record is required" | Provide true or false |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and user permissions |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Table not found | "Table '{tableName}' not found in the file" | Verify table name (case-sensitive) |
| Column not found | "Column '{columnName}' not found in table" | Verify column names (case-sensitive) |
| No matching rows | "No rows found matching the provided keys" | Verify key values exist in first column |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters
- Empty lists for Key Rows or Lookup Column
- Null Get All Record parameter

**Error Message**: `"Drive ID is required"`, `"Key Rows is required"`, `"Lookup Column is required"`

**Resolution**:
1. Verify all required parameters are provided
2. Ensure Key Rows list has at least one value
3. Ensure Lookup Column list has at least one column name
4. Provide boolean value for Get All Record

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Table doesn't exist in the specified file
- Column names don't exist in the table
- No rows match the provided keys (when Get All Record = false)
- File not accessible or doesn't exist
- Insufficient permissions to read file

**Error Message**: `"Table '{tableName}' not found"`, `"Column '{columnName}' not found in table"`

**Resolution**:
1. Verify table name is correct and case-sensitive
2. Check column names are correct and case-sensitive
3. Ensure key values exist in the first column
4. Verify file exists and is accessible
5. Check user has read permissions

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to retrieve column values"`, `"Microsoft Graph API error: {details}"`

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

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite.All permission
3. Check user has read access to the file

## Usage Examples

### Example 1: Get Specific Employee Data

**Scenario**: Retrieve name and department for specific employees

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
Key Rows: ["EMP001", "EMP002"]
Lookup Column: ["Name", "Department"]
Get All Record: false
```

**Output**:
```json
{
  "Result": [
    ["John Doe", "Engineering"],
    ["Jane Smith", "Marketing"]
  ]
}
```

**Result**: Name and department retrieved for EMP001 and EMP002

### Example 2: Get All Records

**Scenario**: Retrieve all employee names and departments

**Input**:
```
Key Rows: ["EMP001"]
Lookup Column: ["Name", "Department"]
Get All Record: true
```

**Output**:
```json
{
  "Result": [
    ["John Doe", "Engineering"],
    ["Jane Smith", "Marketing"],
    ["Bob Johnson", "Sales"]
  ]
}
```

**Result**: All employee names and departments retrieved (Key Rows ignored)

### Example 3: Get Single Column

**Scenario**: Retrieve only names for specific employees

**Input**:
```
Key Rows: ["EMP001", "EMP003"]
Lookup Column: ["Name"]
Get All Record: false
```

**Output**:
```json
{
  "Result": [
    ["John Doe"],
    ["Bob Johnson"]
  ]
}
```

**Result**: Only names retrieved for specified employees

### Example 4: Get Multiple Columns

**Scenario**: Retrieve name, department, and salary

**Input**:
```
Key Rows: ["EMP002"]
Lookup Column: ["Name", "Department", "Salary"]
Get All Record: false
```

**Output**:
```json
{
  "Result": [
    ["Jane Smith", "Marketing", "65000"]
  ]
}
```

**Result**: Multiple column values retrieved for one employee

### Example 5: No Matching Keys

**Scenario**: Search for non-existent employee

**Input**:
```
Key Rows: ["EMP999"]
Lookup Column: ["Name"]
Get All Record: false
```

**Error**: LOGIC_ERROR - "No rows found matching the provided keys"

**Result**: No matching records found

## Business Rules

1. **First Column Search**: Key Rows are matched against the first column of the table
2. **Case-Insensitive Keys**: Key matching is case-insensitive
3. **Case-Sensitive Columns**: Column names are case-sensitive
4. **Multiple Keys**: Can search for multiple keys in one request
5. **Multiple Columns**: Can retrieve multiple columns per row
6. **Get All Override**: When Get All Record = true, all rows are returned regardless of keys
7. **Order Preserved**: Results are returned in the order keys are provided
8. **Partial Matches**: Only exact matches are returned (no partial matching)

## Limitations

1. **First Column Only**: Can only search in the first column of the table
2. **Exact Match**: Does not support partial or wildcard matching
3. **No Sorting**: Results are not sorted (returned in key order or table order)
4. **Shared Drives Only**: This request is for shared drive files only
5. **Table Required**: Cannot perform lookup on regular ranges (must be an Excel table)
6. **No Aggregation**: Cannot perform calculations or aggregations

## Best Practices

### 1. Verify Column Names
Ensure column names match exactly (case-sensitive) before making the request.

### 2. Use Get All Record Wisely
Set Get All Record = false when you only need specific rows to improve performance.

### 3. Validate Key Values
Verify key values exist in the table before performing lookup.

### 4. Limit Lookup Columns
Only request columns you need to reduce data transfer and improve performance.

### 5. Handle Missing Keys
Implement error handling for cases where keys don't exist in the table.

### 6. Cache Table Structure
If performing multiple lookups, cache table structure information.

## Common Use Cases

### 1. Employee Data Lookup
```
Scenario: Get employee details by employee ID
Action: Provide employee IDs as keys, retrieve name and department
Result: Employee information retrieved for display
```

### 2. Product Information Retrieval
```
Scenario: Get product details by product code
Action: Provide product codes as keys, retrieve price and description
Result: Product information available for order processing
```

### 3. Batch Data Retrieval
```
Scenario: Get data for multiple records at once
Action: Provide multiple keys, retrieve multiple columns
Result: Batch data retrieved efficiently
```

### 4. Report Generation
```
Scenario: Generate report with specific data points
Action: Use Get All Record = true to get all data
Result: Complete dataset available for reporting
```

## Related Catalog Requests

- [Table Vlookup](pages/TableVlookup.md) - Single value lookup
- [Get Rows](pages/GetRows.md) - Retrieve all table rows
- [Get Rows for Range](pages/GetRowsForRange.md) - Get data from range
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get File ID
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Get Drive ID

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation, column existence check, key matching
- **Service**: Delegates to WorkbookService for table operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/tables/{tableName}/rows`
- **Processing**: Client-side filtering and column selection
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **ROWS_MATCHED_COUNT**: Number of rows matching keys
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Column not found in table" Error

**Cause**: Column name doesn't match or doesn't exist

**Solution**:
1. Verify column names are spelled correctly (case-sensitive)
2. Check column names match table headers exactly
3. Use "Get Rows" to inspect table structure

### Issue: "No rows found matching the provided keys" Error

**Cause**: Key values don't exist in first column

**Solution**:
1. Verify key values are correct
2. Check first column of table contains the keys
3. Ensure key matching is working (case-insensitive)
4. Use "Get Rows" to see all table data

### Issue: Wrong Columns Returned

**Cause**: Column names don't match expected columns

**Solution**:
1. Verify Lookup Column list contains correct column names
2. Check column names are case-sensitive
3. Inspect table to confirm column names

### Issue: All Rows Returned When Not Expected

**Cause**: Get All Record set to true

**Solution**:
Set Get All Record = false to filter by Key Rows

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Table Vlookup](pages/TableVlookup.md) - Single value lookup
- [Get Rows](pages/GetRows.md) - Get all table rows
- [Get Rows for Range](pages/GetRowsForRange.md) - Get range data

