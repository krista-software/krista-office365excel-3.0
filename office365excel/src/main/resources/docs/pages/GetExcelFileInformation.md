# Get Excel File Information

## Overview

Retrieves information about an Excel file in OneDrive by path and name.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Path | Text | Yes | File path in OneDrive | "/Documents" |
| Name | Text | Yes | File name | "Report.xlsx" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Excel File | Entity(Excel File) | File entity with details |

### Excel File Entity Structure

| Property | Type | Description |
|----------|------|-------------|
| File ID | Text | Unique identifier |
| File Name | Text | Name of the file |
| Web URL | Text | URL to access file |

## Usage Examples

### Example 1: Get File Information

**Input**:
```
Path: "/Documents"
Name: "Sales_Report.xlsx"
```

**Output**:
```json
{
  "Excel File": {
    "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCF",
    "File Name": "Sales_Report.xlsx",
    "Web URL": "https://..."
  }
}
```

## Business Rules

1. **OneDrive Only**: For personal OneDrive files
2. **Exact Match**: Requires exact file name
3. **Path Required**: Must specify folder path

## Related Catalog Requests

- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Search by name only

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/root:/{path}/{name}`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

