# Create Sheet

## Overview

Creates a new worksheet in an Excel file stored in OneDrive.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| New Sheet Name | Text | Yes | Name for the new sheet | "Q1 Data" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sheet | Entity(Excel Sheet) | The newly created sheet entity |

### Excel Sheet Entity Structure

| Property | Type | Description |
|----------|------|-------------|
| Sheet ID | Text | Unique identifier |
| Sheet Name | Text | Name of the sheet |
| Position | Number | Sheet position in workbook |

## Usage Examples

### Example 1: Create New Sheet

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
New Sheet Name: "Q1 Sales"
```

**Output**:
```json
{
  "Sheet": {
    "Sheet ID": "sheet3",
    "Sheet Name": "Q1 Sales",
    "Position": 2
  }
}
```

## Business Rules

1. **Unique Name**: Sheet name must be unique in workbook
2. **Auto-Position**: New sheet added at end
3. **Empty Sheet**: Created sheet has no data

## Limitations

1. **Name Length**: Sheet name limited to 31 characters
2. **Invalid Characters**: Cannot contain: \ / ? * [ ]

## Related Catalog Requests

- [Delete Sheet](pages/DeleteSheet.md) - Delete sheet
- [Get Sheets Information](pages/GetSheetsInformation.md) - List all sheets

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/worksheets`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

