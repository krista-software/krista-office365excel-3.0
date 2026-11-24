# Add Row With Index (Sheet)

## Overview

Inserts rows at a specific position in a worksheet of an Excel file stored in OneDrive.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Rows | List&lt;Text&gt; | No | Comma-separated row data | ["John,Doe,30"] |
| Row Index | Number | Yes | Insert position (1-based) | 5 |
| Rows From Csv | List&lt;File&gt; | No | CSV files | [data.csv] |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sheet Data | List&lt;Composite&gt; | The rows that were inserted |

## Usage Examples

### Example 1: Insert Row at Position 5

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
  "Sheet Data": [
    ["EMP003", "Bob Johnson", "Sales"]
  ]
}
```

**Result**: Row inserted at position 5, existing rows shifted down

## Business Rules

1. **Insert Position**: Rows inserted at specified index
2. **Shift Down**: Existing rows at and below index are shifted down
3. **1-Based Index**: Row Index starts at 1 (not 0)

## Related Catalog Requests

- [Add Rows](pages/AddRows.md) - Append rows (no index)
- [Add Row With Index Async](pages/AddRowWithIndexAsyncSheet.md) - Async version
- [Add Row With Index](pages/AddRowWithIndexShared.md) - Shared drive version

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range/insert`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

