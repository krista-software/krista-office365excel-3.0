# Add Row With Index Async (Sheet)

## Overview

Asynchronously inserts rows at a specific position in a worksheet of an Excel file stored in OneDrive.

## Request Details

- **Area**: Sheet
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Rows | List&lt;Text&gt; | No | Comma-separated row data | ["John,Doe,30"] |
| Rows From Csv | List&lt;File&gt; | No | CSV files | [data.csv] |
| Row Index | Number | Yes | Insert position (1-based) | 5 |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Task ID | Text | Task ID to retrieve results |

## Usage Examples

### Example 1: Insert Row Async

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Employees"
Rows: ["EMP003,Bob Johnson,Sales"]
Row Index: 5
```

**Output**:
```json
{
  "Task ID": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

## Business Rules

1. **Asynchronous**: Runs in background
2. **Insert Position**: Rows inserted at specified index
3. **Shift Down**: Existing rows shifted down

## Related Catalog Requests

- [Get Result](pages/GetResult.md) - Retrieve results
- [Add Row With Index](pages/AddRowWithIndexSheet.md) - Synchronous version
- [Add Row With Index Async](pages/AddRowWithIndexAsync.md) - Shared drive version

## Technical Implementation

### Helper Class
- **Class**: SheetArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range/insert`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

