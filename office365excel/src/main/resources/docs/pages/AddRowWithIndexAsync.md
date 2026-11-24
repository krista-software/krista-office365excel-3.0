# Add Row With Index Async

## Overview

Asynchronously inserts rows at a specific position in a worksheet of an Excel file stored in a shared drive.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
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
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
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

- [Get Add Row With Index Result](pages/GetAddRowWithIndexResult.md) - Retrieve results
- [Add Row With Index](pages/AddRowWithIndexShared.md) - Synchronous version

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range/insert`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

