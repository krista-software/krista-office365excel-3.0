# Vlookup in Table

## Overview

Performs a VLOOKUP operation on an Excel table stored in a shared drive file. Searches for a value in the first column of the table and returns a value from a specified column in the same row.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the table | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Lookup Value | Text | Yes | The value to look up in the first column of the table | "EMP001" |
| Table Name | Text | Yes | The name of the table to perform lookup in | "EmployeeTable" |
| Column Index | Number | Yes | The index of the column to return the result from (1-based) | 2 |
| Exact Match | Boolean | Yes | If true, requires exact match; if false, finds closest match | true |

### Parameter Details

#### Column Index
- **Format**: Positive integer
- **Base**: 1-based index (1 = first column in table)
- **Validation**: Must be ≤ number of columns in table
- **Example**: `2` (second column), `3` (third column)

#### Exact Match
- **True**: Requires exact match of lookup value
- **False**: Finds closest match (approximate match)
- **Recommendation**: Use `true` for most cases

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Vlookup Result | FreeForm | The value found in the specified column |

### Output Example

```json
{
  "Vlookup Result": "John Doe"
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Lookup Value is empty or null | "Lookup Value is required" | Provide a lookup value |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Column Index is null or ≤ 0 | "Column Index must be positive" | Provide valid column index |
| Exact Match is null | "Exact Match is required" | Provide true or false |
| Table not found | "Table '{tableName}' not found" | Verify table name |
| Column Index exceeds table columns | "Column Index exceeds table columns" | Reduce column index |
| Lookup value not found | "Lookup value not found" | Verify value exists |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing or invalid parameters
**Resolution**: Provide all required parameters

### Logic Errors (LOGIC_ERROR)
**Cause**: Lookup value not found or invalid table
**Resolution**: Verify lookup value exists and table is valid

### System Errors (SYSTEM_ERROR)
**Cause**: Microsoft Graph API failures
**Resolution**: Retry operation

### Authorization Errors
**Cause**: Authentication or permission issues
**Resolution**: Re-authenticate, verify permissions

## Usage Examples

### Example 1: Find Employee Name by ID

**Table Structure**:
| Employee ID | Name | Department | Salary |
|-------------|------|------------|--------|
| EMP001 | John Doe | Engineering | 75000 |
| EMP002 | Jane Smith | Marketing | 65000 |

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Lookup Value: "EMP001"
Table Name: "EmployeeTable"
Column Index: 2
Exact Match: true
```

**Output**:
```json
{
  "Vlookup Result": "John Doe"
}
```

### Example 2: Find Department

**Input**:
```
Lookup Value: "EMP002"
Column Index: 3
Exact Match: true
```

**Output**:
```json
{
  "Vlookup Result": "Marketing"
}
```

### Example 3: Value Not Found

**Input**:
```
Lookup Value: "EMP999"
Exact Match: true
```

**Error**: LOGIC_ERROR - "Lookup value not found"

## Business Rules

1. **First Column Search**: Lookup value is searched in the first column of the table
2. **1-Based Index**: Column Index starts at 1 (not 0)
3. **Case-Insensitive**: Lookup matching is case-insensitive
4. **Single Result**: Returns only the first matching value
5. **Formula Values**: Returns calculated values, not formulas

## Limitations

1. **First Column Only**: Can only search in the first column
2. **Single Value**: Returns only one value (first match)
3. **No Wildcards**: Does not support wildcard matching
4. **Shared Drives Only**: For shared drive files only
5. **Table Required**: Must be an Excel table

## Best Practices

### 1. Use Exact Match
Set Exact Match = true for predictable results.

### 2. Verify Column Index
Confirm Column Index is within the table's column count.

### 3. Handle Not Found
Implement error handling for missing lookup values.

## Common Use Cases

### 1. Employee Lookup
```
Scenario: Find employee details by employee ID
Action: VLOOKUP with employee ID
Result: Employee information retrieved
```

### 2. Product Price Lookup
```
Scenario: Get product price by product code
Action: VLOOKUP with product code
Result: Product price retrieved
```

## Related Catalog Requests

- [Table Vlookup](pages/TableVlookup.md) - Personal OneDrive version
- [Sheet Vlookup](pages/SheetVlookup.md) - Range-based VLOOKUP
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md) - Multiple value lookup
- [Get Table Rows](pages/GetTableRows.md) - Get all table data

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Service**: Delegates to WorkbookService for VLOOKUP operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/tables/{tableName}/rows`
- **Processing**: Client-side VLOOKUP logic
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **VLOOKUP_MATCH_FOUND**: Whether match was found

## Troubleshooting

### Issue: "Lookup value not found" Error

**Solution**:
1. Verify lookup value is correct
2. Check value exists in first column of table
3. Try Exact Match = false for approximate match

### Issue: "Column Index exceeds table columns" Error

**Solution**:
1. Count columns in table
2. Ensure Column Index ≤ column count
3. Adjust column index

### Issue: Wrong Value Returned

**Solution**:
1. Remember Column Index is 1-based
2. Count columns carefully from first column
3. Verify which column contains desired data

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)
- [Table Vlookup](pages/TableVlookup.md)
- [Sheet Vlookup](pages/SheetVlookup.md)
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md)

