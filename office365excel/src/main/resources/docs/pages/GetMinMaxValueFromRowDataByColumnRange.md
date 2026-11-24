# Get Min Max Value From Row Data By Column Range

## Overview

Retrieves the minimum and maximum numeric values from a specified range of columns and their corresponding key column values.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sales" |
| Key Column Name | Text | Yes | Column containing key values | "Product" |
| Column Range | Text | Yes | Range to search (e.g., B:D) | "B:D" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Result | FreeForm | Min/max values with corresponding keys |

## Usage Examples

### Example 1: Find Min/Max Sales

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sales"
Key Column Name: "Product"
Column Range: "B:D"
```

**Output**:
```json
{
  "Result": {
    "Minimum": {
      "Value": 100,
      "Key": "Product A",
      "Column": "B"
    },
    "Maximum": {
      "Value": 5000,
      "Key": "Product C",
      "Column": "D"
    }
  }
}
```

## Business Rules

1. **Numeric Only**: Only numeric values considered
2. **Key Association**: Returns key value for min/max
3. **Column Range**: Searches specified column range

## Related Catalog Requests

- [Summation](pages/Summation.md) - Calculate sum of range

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

