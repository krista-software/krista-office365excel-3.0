# Add Rows With List Of Entity (Shared)

## Overview

Adds rows to a worksheet in a shared drive using a list of entity objects, automatically mapping entity properties to columns.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
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
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Employees"
Data: [
  {Employee ID: "EMP001", Name: "John Doe", Department: "Engineering"},
  {Employee ID: "EMP002", Name: "Jane Smith", Department: "Sales"}
]
Exclude Columns: ["Internal ID"]
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
2. **Shared Drives Only**: For shared drive files only

## Related Catalog Requests

- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add rows with text data
- [Add Rows With The List Of Entities](pages/AddRowsWithTheListOfEntities.md) - Sheet version

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range/insert`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

