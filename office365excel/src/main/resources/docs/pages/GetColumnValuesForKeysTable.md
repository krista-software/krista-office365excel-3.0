# Get Column Values for Keys (Table)

## Overview

Retrieves column values for specified key rows and lookup columns from a table.

## Request Details

- **Area**: Table
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Table Name | Text | Yes | Table name | "EmployeeTable" |
| Key Column Name | Text | Yes | Column containing key values | "Employee ID" |
| Keys | List&lt;Text&gt; | Yes | Key values to lookup | ["EMP001", "EMP002"] |
| Lookup Columns | List&lt;Text&gt; | Yes | Columns to retrieve | ["Name", "Department"] |
| Drive Id | Text | No | Drive ID (for shared drives) | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Result | List&lt;Composite&gt; | Retrieved column values for each key |

## Usage Examples

### Example 1: Get Employee Details

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "EmployeeTable"
Key Column Name: "Employee ID"
Keys: ["EMP001", "EMP002"]
Lookup Columns: ["Name", "Department", "Salary"]
```

**Output**:
```json
{
  "Result": [
    {
      "Employee ID": "EMP001",
      "Name": "John Doe",
      "Department": "Engineering",
      "Salary": "75000"
    },
    {
      "Employee ID": "EMP002",
      "Name": "Jane Smith",
      "Department": "Sales",
      "Salary": "80000"
    }
  ]
}
```

## Business Rules

1. **Multi-Key Lookup**: Supports multiple keys
2. **Multi-Column**: Retrieves multiple columns per key
3. **Key Matching**: Matches keys in specified column

## Related Catalog Requests

- [Get Column Values for Keys](pages/GetColumnValuesForKeys.md) - Shared drive version
- [Table Vlookup](pages/TableVlookup.md) - Single value lookup

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/tables/{tableName}/rows`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

