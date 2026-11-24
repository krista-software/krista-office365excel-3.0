# Delete Row With Index

## Overview

Deletes a specific row from a worksheet in an Excel file stored in a shared drive based on the row index.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File ID | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Row Index | Number | Yes | Row position to delete (1-based) | 5 |

> **⚠️ Warning**: This operation is destructive and cannot be undone.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Deleted | Boolean | Indicates if row was successfully deleted |

## Usage Examples

### Example 1: Delete Row

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File ID: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Employees"
Row Index: 5
```

**Output**:
```json
{
  "Deleted": true
}
```

**Result**: Row 5 deleted, rows below shifted up

## Business Rules

1. **Permanent Deletion**: Row and all data permanently deleted
2. **Cannot Undo**: Operation cannot be reversed
3. **1-Based Index**: Row Index starts at 1
4. **Shift Up**: Rows below deleted row shift up

## Limitations

1. **No Undo**: Cannot restore deleted row
2. **Shared Drives Only**: For shared drive files only

## Best Practices

### 1. Backup Data
Create backup before deleting rows.

### 2. Verify Row Index
Double-check row index before deletion.

## Related Catalog Requests

- [Update Row With Index](pages/UpdateRowWithIndex.md)
- [Get Complete Row Info](pages/GetCompleteRowInfo.md)
- [Add Row With Index](pages/AddRowWithIndexShared.md)

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `DELETE /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='A{rowIndex}:{rowIndex}')/delete`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

