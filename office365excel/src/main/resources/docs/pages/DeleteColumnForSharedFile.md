# Delete Column For Shared File

## Overview

Deletes a specific column from an Excel table in a file stored in a shared drive. This operation permanently removes the column and all its data from the table.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive ID | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File ID | Text | Yes | The ID of the Excel file containing the table | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Table Name | Text | Yes | The name of the table containing the column to delete | "EmployeeTable" |
| Column Name | Text | Yes | The name of the column to be deleted | "Temporary Column" |

### Parameter Details

#### Column Name
- **Format**: Name of the column as it appears in the table header
- **Case Sensitivity**: Case-sensitive
- **Validation**: Column must exist in the table
- **Example**: `"Department"`, `"Salary"`, `"Notes"`

> **⚠️ Warning**: This operation is destructive and cannot be undone. All data in the column will be permanently deleted.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Is Deleted | Boolean | Indicates whether the column was successfully deleted |

### Output Example

```json
{
  "Is Deleted": true
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive ID is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File ID is empty or null | "File ID is required" | Provide a valid File ID |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Column Name is empty or null | "Column Name is required" | Provide a valid column name |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and permissions |
| File not found | "File not found or inaccessible" | Verify File ID and permissions |
| Table not found | "Table '{tableName}' not found" | Verify table name (case-sensitive) |
| Column not found | "Column '{columnName}' not found in table" | Verify column name (case-sensitive) |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing or invalid parameters
**Resolution**: Provide all required parameters

### Logic Errors (LOGIC_ERROR)
**Cause**: Table or column doesn't exist, or file not accessible
**Resolution**: Verify table and column names, check file accessibility

### System Errors (SYSTEM_ERROR)
**Cause**: Microsoft Graph API failures
**Resolution**: Retry operation

### Authorization Errors
**Cause**: Authentication or permission issues
**Resolution**: Re-authenticate, verify write permissions

## Usage Examples

### Example 1: Delete Temporary Column

**Table Before**:
| Employee ID | Name | Department | Temporary Column |
|-------------|------|------------|------------------|
| EMP001 | John Doe | Engineering | Data |
| EMP002 | Jane Smith | Marketing | Data |

**Input**:
```
Drive ID: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File ID: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Table Name: "EmployeeTable"
Column Name: "Temporary Column"
```

**Output**:
```json
{
  "Is Deleted": true
}
```

**Table After**:
| Employee ID | Name | Department |
|-------------|------|------------|
| EMP001 | John Doe | Engineering |
| EMP002 | Jane Smith | Marketing |

**Result**: "Temporary Column" and all its data permanently deleted

### Example 2: Delete Notes Column

**Input**:
```
Column Name: "Notes"
```

**Output**:
```json
{
  "Is Deleted": true
}
```

**Result**: "Notes" column removed from table

### Example 3: Column Not Found

**Input**:
```
Column Name: "NonExistentColumn"
```

**Error**: LOGIC_ERROR - "Column 'NonExistentColumn' not found in table"

**Result**: No changes made to table

## Business Rules

1. **Permanent Deletion**: Column and all its data are permanently deleted
2. **Cannot Undo**: Operation cannot be reversed
3. **Case-Sensitive**: Column name matching is case-sensitive
4. **Table Structure**: Table must have at least 2 columns (cannot delete last column)
5. **No Backup**: No automatic backup is created before deletion

## Limitations

1. **Cannot Delete Last Column**: Table must have at least one column remaining
2. **No Undo**: Cannot restore deleted column
3. **Shared Drives Only**: For shared drive files only
4. **Table Required**: Cannot delete columns from regular ranges
5. **No Confirmation**: No confirmation prompt before deletion

## Best Practices

### 1. Verify Column Name
Double-check column name before deletion to avoid deleting wrong column.

### 2. Backup Data
Create a backup of the file before deleting columns.

### 3. Verify Table Structure
Ensure table will have at least one column remaining after deletion.

### 4. Check Dependencies
Verify no formulas or processes depend on the column being deleted.

### 5. Use Caution
This is a destructive operation - use with extreme caution.

## Common Use Cases

### 1. Remove Temporary Columns
```
Scenario: Delete temporary calculation columns no longer needed
Action: Delete column by name
Result: Temporary column removed, table cleaned up
```

### 2. Data Cleanup
```
Scenario: Remove obsolete or deprecated columns
Action: Delete outdated columns
Result: Table structure simplified
```

### 3. Schema Modification
```
Scenario: Modify table structure by removing columns
Action: Delete unwanted columns
Result: Table schema updated
```

## Related Catalog Requests

- [Add Rows to Table](pages/AddRowsToTableShared.md) - Add data to table
- [Get Table Rows](pages/GetTableRows.md) - View table structure
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md) - Read column data
- [Columns Sum](pages/ColumnsSum.md) - Calculate column sums

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation, column existence check
- **Service**: Delegates to WorkbookService for column deletion

### Microsoft Graph API Endpoint
- **Endpoint**: `DELETE /drives/{driveId}/items/{fileId}/workbook/tables/{tableName}/columns/{columnName}`
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **COLUMN_DELETED**: Column name that was deleted

## Troubleshooting

### Issue: "Column not found in table" Error

**Cause**: Column name doesn't match or doesn't exist

**Solution**:
1. Verify column name is spelled correctly (case-sensitive)
2. Check column name matches table header exactly
3. Use "Get Table Rows" to inspect table structure

### Issue: "Cannot delete last column" Error

**Cause**: Attempting to delete the only remaining column

**Solution**:
1. Tables must have at least one column
2. Add a new column before deleting the last one
3. Consider deleting the entire table instead

### Issue: Operation Succeeds But Column Still Visible

**Cause**: Excel file may be cached

**Solution**:
1. Refresh the Excel file
2. Close and reopen the file
3. Wait a few seconds for changes to propagate

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get Table Rows](pages/GetTableRows.md) - View table structure
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Add data to table

