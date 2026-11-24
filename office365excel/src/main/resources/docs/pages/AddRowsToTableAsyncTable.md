# Add Rows To Table Async (Table Area)

## Overview

Asynchronously adds rows to a table at a specific index position.

## Request Details

- **Area**: Table
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Drive Id | Text | No | Drive ID (for shared drives) | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| Table Name | Text | Yes | Table name | "EmployeeTable" |
| Rows | List&lt;Text&gt; | No | Comma-separated row data | ["John,Doe,30"] |
| Rows From Csv | List&lt;File&gt; | No | CSV files | [data.csv] |
| Row Index | Number | Yes | Insert position (1-based) | 5 |
| Separator | Text | No | CSV delimiter (default: comma) | "," |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| TaskId | Text | Task ID to retrieve results |

## Usage Examples

### Example 1: Insert Rows Async

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Rows: ["EMP003,Bob Johnson,Sales"]
Row Index: 5
```

**Output**:
```json
{
  "TaskId": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

## Business Rules

1. **Asynchronous**: Runs in background
2. **Insert Position**: Rows inserted at specified index
3. **Batch Processing**: Supports large datasets

## Related Catalog Requests

- [Get Add Rows To Table Result](pages/GetAddRowsToTableResultTable.md) - Retrieve results
- [Add Rows To Table](pages/AddRowsToTableTable.md) - Synchronous version

## Technical Implementation

### Helper Class
- **Class**: TableArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/tables/{tableName}/rows/add`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

