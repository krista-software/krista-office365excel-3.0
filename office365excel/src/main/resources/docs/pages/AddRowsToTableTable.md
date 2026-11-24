# Add Rows To Table (Table Area)

## Overview

Adds rows to a table at a specific index position.

## Request Details

- **Area**: Table
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Drive ID | Text | No | Drive ID (for shared drives) | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| Table Name | Text | Yes | Table name | "EmployeeTable" |
| Rows | List&lt;Text&gt; | No | Comma-separated row data | ["John,Doe,30"] |
| Rows From Csv | List&lt;File&gt; | No | CSV files | [data.csv] |
| Index | Number | Yes | Insert position (1-based) | 5 |
| Separator | Text | No | CSV delimiter (default: comma) | "," |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Total Rows | List&lt;Composite&gt; | The rows that were inserted |

## Usage Examples

### Example 1: Insert Rows at Position

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Rows: ["EMP003,Bob Johnson,Sales"]
Index: 5
```

**Output**:
```json
{
  "Total Rows": [
    ["EMP003", "Bob Johnson", "Sales"]
  ]
}
```

## Business Rules

1. **Insert Position**: Rows inserted at specified index
2. **Shift Down**: Existing rows shifted down
3. **1-Based Index**: Index starts at 1

## Related Catalog Requests

- [Add Rows To Table Async](pages/AddRowsToTableAsync.md) - Async version
- [Add Rows](pages/AddRows.md) - Append to end

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/tables/{tableName}/rows/add`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

