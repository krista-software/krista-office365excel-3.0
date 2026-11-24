# Add Rows Async (Table)

## Overview

Asynchronously adds rows to a table in an Excel file.

## Request Details

- **Area**: Table
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Table Name | Text | Yes | Table name | "EmployeeTable" |
| Rows | List&lt;Text&gt; | No | Comma-separated row data | ["John,Doe,30"] |
| Rows From Csv | List&lt;File&gt; | No | CSV files | [data.csv] |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Task ID | Text | Task ID to retrieve results |

## Usage Examples

### Example 1: Add Rows Async

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Rows: ["EMP003,Bob Johnson,Sales"]
```

**Output**:
```json
{
  "Task ID": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

## Business Rules

1. **Asynchronous**: Runs in background
2. **Append**: Rows added to end of table
3. **Task ID Required**: Use Task ID to retrieve results

## Related Catalog Requests

- [Get Add Rows Result](pages/GetAddRowsResult.md) - Retrieve results
- [Add Rows](pages/AddRows.md) - Synchronous version

## Technical Implementation

### Helper Class
- **Class**: TableArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/tables/{tableName}/rows`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

