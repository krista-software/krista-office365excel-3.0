# Get Column Values For Keys In Sheet

## Overview

Retrieves column values based on key rows and columns from a specified range in a worksheet.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Sheet Name | Text | Yes | Sheet name | "Employees" |
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Range | Text | Yes | Cell range (e.g., A1:D10) | "A1:D100" |
| Key Rows | List&lt;Text&gt; | Yes | Key values to lookup | ["EMP001", "EMP002"] |
| Key Columns | List&lt;Text&gt; | Yes | Columns to retrieve | ["Name", "Department"] |
| Get All Records | Boolean | Yes | true = all matches, false = first match | true |
| Column Number | Number | No | Key column number (1-based) | 1 |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Data | List&lt;Composite&gt; | Retrieved column values for each key |

## Usage Examples

### Example 1: Get Employee Details

**Input**:
```
Sheet Name: "Employees"
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Range: "A1:D100"
Key Rows: ["EMP001", "EMP002"]
Key Columns: ["Name", "Department", "Salary"]
Get All Records: true
Column Number: 1
```

**Output**:
```json
{
  "Data": [
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
3. **Range-Based**: Searches within specified range

## Related Catalog Requests

- [Get Column Values for Keys](pages/GetColumnValuesForKeys.md) - Shared drive version
- [Get Column Values for Keys](pages/GetColumnValuesForKeysTable.md) - Table version

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

