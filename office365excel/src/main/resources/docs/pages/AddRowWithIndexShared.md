# Add Row With Index (Shared)

## Overview

Inserts rows at a specific position in a worksheet of an Excel file stored in a shared drive. Unlike append operations, this allows inserting rows at any position.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the sheet | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | The name of the sheet where the row will be inserted | "Sheet1" |
| Rows | List&lt;Text&gt; | No | List of comma-separated row data | ["John,Doe,30"] |
| Row Index | Number | Yes | The position where the row will be inserted (1-based) | 5 |
| Rows From Csv | List&lt;File&gt; | No | CSV files containing rows to insert | [data.csv] |

> **📝 Note**: At least one of "Rows" or "Rows From Csv" must be provided.

### Parameter Details

#### Row Index
- **Format**: Positive integer
- **Base**: 1-based index (1 = first row)
- **Behavior**: Existing rows at and below this index are shifted down
- **Example**: `5` (insert at row 5, shifting current row 5 down)

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sheet Data | List&lt;Composite&gt; | The rows that were inserted |

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive Id is required" | Provide a valid Drive Id |
| File Id is empty or null | "File Id is required" | Provide a valid File Id |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Row Index is null or ≤ 0 | "Row Index must be positive" | Provide valid row index |
| Both Rows and Rows From Csv are empty | "Row data not provided" | Provide at least one data source |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing required parameters or no row data
**Resolution**: Provide all required parameters and row data

### Logic Errors (LOGIC_ERROR)
**Cause**: Sheet doesn't exist or file not accessible
**Resolution**: Verify sheet name and file accessibility

## Usage Examples

### Example 1: Insert Row at Position 5

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Employees"
Rows: ["EMP003,Bob Johnson,Sales"]
Row Index: 5
Rows From Csv: null
```

**Output**:
```json
{
  "Sheet Data": [
    ["EMP003", "Bob Johnson", "Sales"]
  ]
}
```

**Result**: Row inserted at position 5, existing rows shifted down

### Example 2: Insert Multiple Rows

**Input**:
```
Rows: ["Row1,Data1", "Row2,Data2"]
Row Index: 3
```

**Result**: Two rows inserted starting at position 3

### Example 3: Insert from CSV

**Input**:
```
Rows: null
Rows From Csv: [new_employees.csv]
Row Index: 10
```

**Result**: CSV rows inserted starting at position 10

## Business Rules

1. **Insert Position**: Rows inserted at specified index
2. **Shift Down**: Existing rows at and below index are shifted down
3. **1-Based Index**: Row Index starts at 1 (not 0)
4. **Order Preserved**: Multiple rows inserted in order provided
5. **No Overwrite**: Existing data is never overwritten

## Limitations

1. **Shared Drives Only**: For shared drive files only
2. **No Formatting**: Inserted rows have no formatting
3. **Batch Size**: Large batches may hit API limits

## Best Practices

### 1. Verify Row Index
Ensure Row Index is within valid range for the sheet.

### 2. Use for Specific Positioning
Use this when row position matters; use append for adding to end.

### 3. Handle Shifts
Be aware that existing rows will shift down.

## Common Use Cases

### 1. Insert Header Row
```
Scenario: Insert header row at top of data
Action: Insert row at index 1
Result: Header inserted, data shifted down
```

### 2. Insert Data at Specific Position
```
Scenario: Insert new record between existing records
Action: Insert row at specific index
Result: New record inserted, others shifted
```

## Related Catalog Requests

- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Append rows (no index)
- [Add Row With Index Async](pages/AddRowWithIndexAsync.md) - Async version
- [Add Row With Index](pages/AddRowWithIndex.md) - Sheet Area version

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')/insert`
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **ROWS_INSERTED_COUNT**: Number of rows inserted
- **INSERT_INDEX**: Position where rows were inserted

## Troubleshooting

### Issue: Rows Inserted at Wrong Position

**Cause**: Row Index is 1-based, not 0-based

**Solution**: Remember Row Index 1 = first row, not row 0

### Issue: Existing Data Shifted

**Cause**: This is expected behavior

**Solution**: This operation shifts existing rows down by design

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)
- [Add Rows to Sheet](pages/AddRowsToSheet.md)
- [Add Row With Index Async](pages/AddRowWithIndexAsync.md)

