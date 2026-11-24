# Get Sheets Information

## Overview

Returns names and IDs of all sheets defined in an Excel file.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sheets | List&lt;Entity(Excel Sheet)&gt; | List of sheet entities |

### Excel Sheet Entity Structure

| Property | Type | Description |
|----------|------|-------------|
| Sheet ID | Text | Unique identifier |
| Sheet Name | Text | Name of the sheet |
| Position | Number | Sheet position in workbook |

## Usage Examples

### Example 1: Get All Sheets

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
```

**Output**:
```json
{
  "Sheets": [
    {
      "Sheet ID": "sheet1",
      "Sheet Name": "Sheet1",
      "Position": 0
    },
    {
      "Sheet ID": "sheet2",
      "Sheet Name": "Data",
      "Position": 1
    }
  ]
}
```

## Business Rules

1. **All Sheets**: Returns all sheets in workbook
2. **Ordered**: Sheets returned in order of position

## Related Catalog Requests

- [Create Sheet](pages/CreateSheet.md) - Create new sheet
- [Delete Sheet](pages/DeleteSheet.md) - Delete sheet

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/worksheets`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

