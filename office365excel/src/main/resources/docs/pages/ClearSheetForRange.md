# Clear Sheet For Range

## Overview

Clears all data from a specified range of cells in a worksheet.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
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
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
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

### Example 2: Clear Entire Column

**Input**:
```
Cell Range: "A:A"
```

**Result**: Entire column A cleared

## Business Rules

1. **Permanent**: Data cannot be recovered
2. **Range Notation**: Supports Excel range notation
3. **Preserves Formatting**: Only clears values, not formatting

## Limitations

1. **No Undo**: Cannot restore cleared data
2. **OneDrive Only**: For personal OneDrive files

## Best Practices

### 1. Backup Data
Create backup before clearing ranges.

### 2. Verify Range
Double-check range before clearing.

## Related Catalog Requests

- [Get Rows for Range](pages/GetRowsForRange.md) - View range data before clearing

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')/clear`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

