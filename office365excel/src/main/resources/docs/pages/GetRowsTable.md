# Get Rows (Table)

## Overview

Returns all rows from a specified table in an Excel workbook.

## Request Details

- **Area**: Table
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Workbook ID | Text | Yes | Workbook/File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Table Name | Text | Yes | Table name | "EmployeeTable" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Table Rows | List&lt;Composite&gt; | All rows from the table |

## Usage Examples

### Example 1: Get All Table Rows

**Input**:
```
Workbook ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
```

**Output**:
```json
{
  "Table Rows": [
    ["EMP001", "John Doe", "Engineering"],
    ["EMP002", "Jane Smith", "Sales"],
    ["EMP003", "Bob Johnson", "Marketing"]
  ]
}
```

## Business Rules

1. **All Rows**: Returns all rows in table
2. **Excludes Headers**: Header row not included
3. **Ordered**: Rows returned in table order

## Related Catalog Requests

- [Get Table Rows](pages/GetTableRows.md) - Shared drive version
- [Add Rows](pages/AddRows.md) - Add rows to table

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{workbookId}/workbook/tables/{tableName}/rows`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

