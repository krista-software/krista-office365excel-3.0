# Summation (Table)

## Overview

Calculates the sum of numeric values in specified columns within a table.

## Request Details

- **Area**: Table
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Table Name | Text | Yes | Table name | "SalesTable" |
| Columns | List&lt;Text&gt; | Yes | Column names to sum | ["Q1 Sales", "Q2 Sales"] |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sum | Number | The calculated sum of all specified columns |

## Usage Examples

### Example 1: Sum Sales Columns

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Table Name: "SalesTable"
Columns: ["Q1 Sales", "Q2 Sales", "Q3 Sales"]
```

**Output**:
```json
{
  "Sum": 450000.75
}
```

## Business Rules

1. **Multiple Columns**: Sums all specified columns
2. **Numeric Only**: Only numeric values included
3. **Ignores Text**: Text values ignored

## Related Catalog Requests

- [Summation](pages/Summation.md) - Sheet version
- [Columns Sum](pages/ColumnsSum.md) - Shared drive version

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/functions/sum`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

