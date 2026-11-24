# Sheet Vlookup

## Overview

Performs a VLOOKUP operation on a specified range within a worksheet of an Excel file stored in a shared drive. Searches for a value in the first column of the range and returns a value from a specified column in the same row.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the sheet | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Lookup Value | Text | Yes | The value to look up in the first column of the range | "EMP001" |
| Sheet Name | Text | Yes | The name of the sheet containing the lookup range | "Employees" |
| Range | Text | Yes | The Excel range to search for the lookup value | "A1:D100" |
| Column Index | Number | Yes | The index of the column to return the result from (1-based) | 2 |
| Exact Match | Boolean | Yes | If true, requires exact match; if false, finds closest match | true |

### Parameter Details

#### Lookup Value
- **Format**: Text or numeric value to search for
- **Search Location**: First column of the specified range
- **Case Sensitivity**: Case-insensitive matching
- **Example**: `"EMP001"`, `"Product A"`, `"12345"`

#### Range
- **Format**: Excel range notation (e.g., A1:D100)
- **Pattern**: `[A-Z]+[0-9]+:[A-Z]+[0-9]+`
- **Must Include**: At least 2 columns (lookup column + return column)
- **Example**: `A1:D100`, `B5:F50`

#### Column Index
- **Format**: Positive integer
- **Base**: 1-based index (1 = first column in range)
- **Validation**: Must be ≤ number of columns in range
- **Example**: `2` (second column), `3` (third column)

#### Exact Match
- **True**: Requires exact match of lookup value
- **False**: Finds closest match (approximate match)
- **Recommendation**: Use `true` for most cases
- **Example**: `true`, `false`

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
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Range is empty or null | "Range is required" | Provide a valid range |
| Column Index is null or ≤ 0 | "Column Index must be positive" | Provide valid column index |
| Exact Match is null | "Exact Match is required" | Provide true or false |
| Invalid range format | "Invalid range format" | Use format: A1:D100 |
| Column Index exceeds range | "Column Index exceeds range columns" | Reduce column index |
| Lookup value not found | "Lookup value not found" | Verify value exists in first column |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing or invalid parameters
**Resolution**: Provide all required parameters with valid formats

### Logic Errors (LOGIC_ERROR)
**Cause**: Lookup value not found or invalid range
**Resolution**: Verify lookup value exists and range is valid

### System Errors (SYSTEM_ERROR)
**Cause**: Microsoft Graph API failures
**Resolution**: Retry operation

### Authorization Errors
**Cause**: Authentication or permission issues
**Resolution**: Re-authenticate, verify permissions

## Usage Examples

### Example 1: Find Employee Name by ID

**Range Data**:
| A (Col 1) | B (Col 2) | C (Col 3) | D (Col 4) |
|-----------|-----------|-----------|-----------|
| EMP001 | John Doe | Engineering | 75000 |
| EMP002 | Jane Smith | Marketing | 65000 |
| EMP003 | Bob Johnson | Sales | 70000 |

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Lookup Value: "EMP001"
Sheet Name: "Employees"
Range: "A1:D100"
Column Index: 2
Exact Match: true
```

**Output**:
```json
{
  "Vlookup Result": "John Doe"
}
```

**Result**: Found "John Doe" in column 2 for employee EMP001

### Example 2: Find Department by Employee ID

**Input**:
```
Lookup Value: "EMP002"
Range: "A1:D100"
Column Index: 3
Exact Match: true
```

**Output**:
```json
{
  "Vlookup Result": "Marketing"
}
```

**Result**: Found "Marketing" in column 3 for employee EMP002

### Example 3: Find Salary by Employee ID

**Input**:
```
Lookup Value: "EMP003"
Range: "A1:D100"
Column Index: 4
Exact Match: true
```

**Output**:
```json
{
  "Vlookup Result": "70000"
}
```

**Result**: Found "70000" in column 4 for employee EMP003

### Example 4: Approximate Match

**Input**:
```
Lookup Value: "EMP001.5"
Range: "A1:D100"
Column Index: 2
Exact Match: false
```

**Output**:
```json
{
  "Vlookup Result": "John Doe"
}
```

**Result**: Found closest match (EMP001)

### Example 5: Value Not Found

**Input**:
```
Lookup Value: "EMP999"
Exact Match: true
```

**Error**: LOGIC_ERROR - "Lookup value not found"

**Result**: No matching value in first column

## Business Rules

1. **First Column Search**: Lookup value is searched in the first column of the range
2. **1-Based Index**: Column Index starts at 1 (not 0)
3. **Case-Insensitive**: Lookup matching is case-insensitive
4. **Exact Match Recommended**: Use Exact Match = true for predictable results
5. **Single Result**: Returns only the first matching value
6. **Formula Values**: Returns calculated values, not formulas

## Limitations

1. **First Column Only**: Can only search in the first column of the range
2. **Single Value**: Returns only one value (first match)
3. **No Wildcards**: Does not support wildcard matching
4. **Shared Drives Only**: For shared drive files only
5. **Range Required**: Cannot search entire sheet (must specify range)

## Best Practices

### 1. Use Exact Match
Set Exact Match = true for predictable, reliable results.

### 2. Validate Range
Ensure range includes both lookup column and return column.

### 3. Verify Column Index
Confirm Column Index is within the range's column count.

### 4. Handle Not Found
Implement error handling for cases where lookup value doesn't exist.

### 5. Sort Data for Approximate Match
If using Exact Match = false, ensure data is sorted ascending.

## Common Use Cases

### 1. Employee Lookup
```
Scenario: Find employee details by employee ID
Action: VLOOKUP with employee ID in first column
Result: Employee name, department, or salary retrieved
```

### 2. Product Price Lookup
```
Scenario: Get product price by product code
Action: VLOOKUP with product code
Result: Product price retrieved
```

### 3. Customer Information Retrieval
```
Scenario: Find customer details by customer ID
Action: VLOOKUP with customer ID
Result: Customer name, email, or phone retrieved
```

## Related Catalog Requests

- [Table Vlookup](pages/TableVlookup.md) - VLOOKUP in Excel table
- [Vlookup in Table](pages/VlookupInTable.md) - Table-based VLOOKUP
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md) - Multiple value lookup
- [Get Rows for Range](pages/GetRowsForRange.md) - Get range data

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Service**: Delegates to WorksheetService for VLOOKUP operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')`
- **Processing**: Client-side VLOOKUP logic on retrieved range data
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **VLOOKUP_MATCH_FOUND**: Whether match was found

## Troubleshooting

### Issue: "Lookup value not found" Error

**Cause**: Value doesn't exist in first column of range

**Solution**:
1. Verify lookup value is correct
2. Check value exists in first column of range
3. Ensure range includes the row with the value
4. Try Exact Match = false for approximate match

### Issue: "Column Index exceeds range columns" Error

**Cause**: Column Index is larger than number of columns in range

**Solution**:
1. Count columns in range (e.g., A1:D100 has 4 columns)
2. Ensure Column Index ≤ column count
3. Adjust range or column index

### Issue: Wrong Value Returned

**Cause**: Column Index pointing to wrong column

**Solution**:
1. Remember Column Index is 1-based (1 = first column)
2. Count columns carefully from first column of range
3. Verify which column contains desired data

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)
- [Table Vlookup](pages/TableVlookup.md)
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md)

