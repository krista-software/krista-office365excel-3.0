# Add Formula To Cell

## Overview

Adds an Excel formula to a specific cell in a worksheet.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Cell ID | Text | Yes | Cell address (e.g., A1) | "C1" |
| Formula | Text | Yes | Excel formula | "=SUM(A1:B1)" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell Information | FreeForm | Updated cell information |

## Usage Examples

### Example 1: Add SUM Formula

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sheet1"
Cell ID: "C1"
Formula: "=SUM(A1:B1)"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "C1",
    "Formula": "=SUM(A1:B1)",
    "Value": 150
  }
}
```

### Example 2: Add VLOOKUP Formula

**Input**:
```
Cell ID: "D2"
Formula: "=VLOOKUP(A2,A1:C10,3,FALSE)"
```

**Result**: VLOOKUP formula added to cell D2

## Business Rules

1. **Formula Syntax**: Must start with =
2. **Auto-Calculate**: Formula calculated automatically
3. **Overwrites**: Replaces existing cell content

## Related Catalog Requests

- [Execute Formula](pages/ExecuteFormula.md) - Execute without modifying sheet
- [Add Data To Cell](pages/AddDataToCell.md) - Add data or formula

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

