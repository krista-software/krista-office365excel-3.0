# Add Rows With The List Of Entities

## Overview

Adds rows to a worksheet using a list of entity objects, automatically mapping entity properties to columns.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Employees" |
| Data | List&lt;Entity&gt; | Yes | List of entity objects | [Employee entities] |
| Exclude Columns | List&lt;Text&gt; | No | Columns to exclude | ["Internal ID"] |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Final Data | List&lt;Composite&gt; | The rows that were added |

## Usage Examples

### Example 1: Add Employee Entities

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Employees"
Data: [
  {Employee ID: "EMP001", Name: "John Doe", Department: "Engineering"},
  {Employee ID: "EMP002", Name: "Jane Smith", Department: "Sales"}
]
Exclude Columns: ["Internal ID", "Created Date"]
```

**Output**:
```json
{
  "Final Data": [
    ["EMP001", "John Doe", "Engineering"],
    ["EMP002", "Jane Smith", "Sales"]
  ]
}
```

## Business Rules

1. **Entity Mapping**: Automatically maps entity properties to columns
2. **Exclude Columns**: Can exclude specific properties
3. **Header Detection**: Detects if sheet has headers

## Limitations

1. **Entity Type**: All entities must be same type
2. **OneDrive Only**: For personal OneDrive files

## Related Catalog Requests

- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add rows with text data
- [Add Rows](pages/AddRows.md) - Add rows to table

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

