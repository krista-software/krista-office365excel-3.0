# Delete Column With Index

## Overview

Deletes a specific column from a worksheet in an Excel file stored in a shared drive. This operation permanently removes the column and all its data.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | The ID of the Excel file containing the sheet | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| Sheet Name | Text | Yes | The name of the sheet containing the column to delete | "Sheet1" |
| Column Name | Text | Yes | The name of the column to be deleted | "Temporary Column" |

> **⚠️ Warning**: This operation is destructive and cannot be undone.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Deleted | Boolean | Indicates whether the column was successfully deleted |

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| File Id is empty or null | "File Id is required" | Provide a valid File Id |
| Drive Id is empty or null | "Drive Id is required" | Provide a valid Drive Id |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Column Name is empty or null | "Column Name is required" | Provide a valid column name |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing required parameters
**Resolution**: Provide all required parameters

### Logic Errors (LOGIC_ERROR)
**Cause**: Sheet or column doesn't exist
**Resolution**: Verify sheet and column names

## Usage Examples

### Example 1: Delete Column

**Input**:
```
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
Sheet Name: "Sheet1"
Column Name: "Temporary Column"
```

**Output**:
```json
{
  "Deleted": true
}
```

## Business Rules

1. **Permanent Deletion**: Column and all data permanently deleted
2. **Cannot Undo**: Operation cannot be reversed
3. **Case-Sensitive**: Column name matching is case-sensitive

## Limitations

1. **No Undo**: Cannot restore deleted column
2. **Shared Drives Only**: For shared drive files only

## Best Practices

### 1. Backup Data
Create backup before deleting columns.

### 2. Verify Column Name
Double-check column name before deletion.

## Common Use Cases

### 1. Remove Temporary Columns
```
Scenario: Delete temporary calculation columns
Action: Delete column by name
Result: Column removed
```

## Related Catalog Requests

- [Delete Column For Shared File](pages/DeleteColumnForSharedFile.md) - Table version
- [Get Rows for Range](pages/GetRowsForRange.md) - View sheet data

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `DELETE /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/columns/{columnName}`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

