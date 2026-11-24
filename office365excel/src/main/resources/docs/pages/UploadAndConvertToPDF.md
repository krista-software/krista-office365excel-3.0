# Upload And Convert To PDF

## Overview

Uploads an Excel file to OneDrive and converts it to PDF format.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Excel File | File | Yes | Excel file to upload and convert | report.xlsx |
| File Name | Text | Yes | Name for the uploaded file | "Q1_Report.xlsx" |
| Folder Path | Text | No | OneDrive folder path | "/Reports" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| PDF File | File | The converted PDF file |
| File Id | Text | ID of the uploaded Excel file in OneDrive |

## Usage Examples

### Example 1: Upload and Convert

**Input**:
```
Excel File: [report.xlsx binary data]
File Name: "Q1_Sales_Report.xlsx"
Folder Path: "/Reports"
```

**Output**:
```json
{
  "PDF File": [PDF binary data],
  "File Id": "01C8MOQGI3YRBDXAQERECWE3W4XCF"
}
```

## Business Rules

1. **Upload First**: File uploaded to OneDrive before conversion
2. **Preserves Formatting**: Original formatting maintained in PDF
3. **All Sheets**: All worksheets included in PDF

## Limitations

1. **File Size**: Large files may take longer
2. **OneDrive Only**: Uploads to personal OneDrive

## Related Catalog Requests

- [Convert To PDF](pages/ConvertToPDF.md) - Convert existing file

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `PUT /me/drive/root:/{path}/{name}:/content`
- **Endpoint**: `GET /me/drive/items/{fileId}/content?format=pdf`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

