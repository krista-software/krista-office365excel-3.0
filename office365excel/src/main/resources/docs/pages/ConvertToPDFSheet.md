# Convert To PDF (Sheet)

## Overview

Converts an Excel file from OneDrive to PDF format while maintaining the same layout.

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
| PDF File | File | The converted PDF file |

## Usage Examples

### Example 1: Convert to PDF

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
```

**Output**:
```
PDF File: [Binary PDF file data]
```

## Business Rules

1. **Preserves Layout**: Original formatting maintained
2. **All Sheets**: All worksheets included in PDF
3. **Read-Only**: Original Excel file not modified

## Related Catalog Requests

- [Convert To PDF](pages/ConvertToPDF.md) - Shared drive version
- [Upload And Convert To PDF](pages/UploadAndConvertToPDF.md) - Upload and convert

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/content?format=pdf`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

