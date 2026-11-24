# Create Table for Range

## Overview

Creates an Excel table from a specified range of cells in a worksheet.

## Request Details

- **Area**: Table
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Range | Text | Yes | Cell range (e.g., A1:D10) | "A1:D10" |
| Drive Id | Text | No | Drive ID (for shared drives) | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Table | Entity(Table) | The newly created table entity |

### Table Entity Structure

| Property | Type | Description |
|----------|------|-------------|
| Table ID | Text | Unique identifier |
| Table Name | Text | Name of the table |
| Sheet Name | Text | Worksheet containing the table |

## Usage Examples

### Example 1: Create Table

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sheet1"
Range: "A1:D10"
```

**Output**:
```json
{
  "Table": {
    "Table ID": "table1",
    "Table Name": "Table1",
    "Sheet Name": "Sheet1"
  }
}
```

## Business Rules

1. **Header Row**: First row of range becomes table headers
2. **Auto-Name**: Table automatically named if not specified
3. **Structured**: Converts range to structured table

## Limitations

1. **Range Must Have Data**: Range must contain data
2. **No Overlapping**: Cannot overlap existing tables

## Related Catalog Requests

- [Get Rows](pages/GetRowsTable.md) - Get table rows
- [Add Rows](pages/AddRows.md) - Add rows to table

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/tables/add`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

