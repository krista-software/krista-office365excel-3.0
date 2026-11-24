# Add Rows To Table In Batch Async

## Overview

Adds multiple rows of data to an Excel table in a single asynchronous batch operation with progress tracking, allowing efficient bulk insertion without blocking the client.

## Request Details

- **Area**: Table
- **Type**: CHANGE_SYSTEM
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
| Separator | Text | Yes | CSV delimiter | "," |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| TaskId | Text | Task ID to track progress and retrieve results |

## Usage Examples

### Example 1: Batch Insert with Progress

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Rows From Csv: [employees.csv]
Row Index: 1
Separator: ","
```

**Output**:
```json
{
  "TaskId": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

## Business Rules

1. **Batch Processing**: Processes large datasets efficiently
2. **Progress Tracking**: Sends progress updates (~every 20%)
3. **Asynchronous**: Runs in background

## Related Catalog Requests

- [Add Rows Async To Table Progress Update](pages/AddRowsAsyncToTableProgressUpdate.md) - Track progress
- [Add Rows Async To Table Result Available](pages/AddRowsAsyncToTableResultAvailable.md) - Get final results

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

