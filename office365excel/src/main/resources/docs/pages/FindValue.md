# Find Value

## Overview

Searches for a specific value in a worksheet and returns the cell IDs where the value is found.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Lookup Value | Text | Yes | Value to search for | "John Doe" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell IDs | List&lt;Text&gt; | List of cell addresses where value was found |

## Usage Examples

### Example 1: Find Value

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Employees"
Lookup Value: "John Doe"
```

**Output**:
```json
{
  "Cell IDs": ["B2", "D5"]
}
```

## Business Rules

1. **Case-Insensitive**: Search is case-insensitive
2. **Exact Match**: Only exact matches returned
3. **All Occurrences**: Returns all cells containing the value

## Related Catalog Requests

- [Search](pages/Search.md) - Shared drive version
- [Get Cell Value](pages/GetCellValue.md) - Read cell values

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/usedRange`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

