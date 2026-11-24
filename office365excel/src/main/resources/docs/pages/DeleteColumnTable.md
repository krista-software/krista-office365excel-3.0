# Delete Column (Table)

## Overview

Deletes a specific column from a table in an Excel spreadsheet.

## Request Details

- **Area**: Table
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Table Name | Text | Yes | Table name | "EmployeeTable" |
| Column Name | Text | Yes | Column name to delete | "Temporary Column" |

> **⚠️ Warning**: This operation is destructive and cannot be undone.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Is Deleted | Boolean | Indicates if column was successfully deleted |

## Usage Examples

### Example 1: Delete Column

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Column Name: "Temporary Column"
```

**Output**:
```json
{
  "Is Deleted": true
}
```

## Business Rules

1. **Permanent Deletion**: Column and all data permanently deleted
2. **Cannot Undo**: Operation cannot be reversed
3. **Case-Sensitive**: Column name matching is case-sensitive

## Limitations

1. **No Undo**: Cannot restore deleted column
2. **Table Only**: Only works with Excel tables

## Best Practices

### 1. Backup Data
Create backup before deleting columns.

### 2. Verify Column Name
Double-check column name before deletion.

## Related Catalog Requests

- [Delete Column For Shared File](pages/DeleteColumnForSharedFile.md) - Shared drive version
- [Delete Column With Index](pages/DeleteColumnWithIndex.md) - Sheet version

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `DELETE /me/drive/items/{fileId}/workbook/tables/{tableName}/columns/{columnName}`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

