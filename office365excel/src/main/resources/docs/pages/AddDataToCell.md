# Add Data To Cell

## Overview

Adds or updates data in a specific cell of a worksheet in an Excel file stored in OneDrive.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Cell ID | Text | Yes | Cell address (e.g., A1, B2) | "A1" |
| Data | Text | Yes | Data to add to cell | "Hello World" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell Information | FreeForm | Updated cell information |

## Usage Examples

### Example 1: Add Text to Cell

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sheet1"
Cell ID: "A1"
Data: "Employee Name"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "A1",
    "Value": "Employee Name"
  }
}
```

### Example 2: Add Formula

**Input**:
```
Cell ID: "C1"
Data: "=SUM(A1:B1)"
```

**Result**: Formula added to cell C1

## Business Rules

1. **Overwrites**: Replaces existing cell value
2. **Formula Support**: Can write formulas (start with =)
3. **Clear Cell**: Use empty string to clear cell

## Related Catalog Requests

- [Add Data to Provided Cell](pages/AddDataToProvidedCell.md) - Shared drive version
- [Get Cell Value](pages/GetCellValue.md) - Read cell value
- [Add Multiple Cell Data](pages/AddMultipleCellData.md) - Update multiple cells

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `PATCH /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{cellId}')`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

