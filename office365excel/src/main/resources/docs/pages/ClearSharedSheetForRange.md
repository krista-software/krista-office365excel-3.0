# Clear Shared Sheet For Range

## Overview

Clears all data from a specified range of cells in a worksheet stored in a shared drive.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Cell Range | Text | Yes | Range to clear (e.g., A1:D10) | "A1:D10" |

> **⚠️ Warning**: This operation permanently clears data and cannot be undone.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cleared | Boolean | Indicates if range was successfully cleared |

## Usage Examples

### Example 1: Clear Range

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Sheet1"
Cell Range: "A1:D10"
```

**Output**:
```json
{
  "Cleared": true
}
```

**Result**: All data in range A1:D10 cleared

## Business Rules

1. **Permanent**: Data cannot be recovered
2. **Range Notation**: Supports Excel range notation
3. **Preserves Formatting**: Only clears values, not formatting

## Limitations

1. **No Undo**: Cannot restore cleared data
2. **Shared Drives Only**: For shared drive files only

## Best Practices

### 1. Backup Data
Create backup before clearing ranges.

### 2. Verify Range
Double-check range before clearing.

## Related Catalog Requests

- [Clear Sheet For Range](pages/ClearSheetForRange.md) - Sheet version
- [Get Rows for Range](pages/GetRowsForRange.md) - View range data before clearing

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')/clear`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

