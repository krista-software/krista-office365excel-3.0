# Summation

## Overview

Calculates the sum of numeric values within a specified range in an Excel worksheet.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | Sheet name | "Sales" |
| Range | Text | Yes | Cell range to sum (e.g., A1:B10) | "A1:B10" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sum | Number | The calculated sum of all numeric values |

## Usage Examples

### Example 1: Sum Sales Data

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sales"
Range: "B2:B100"
```

**Output**:
```json
{
  "Sum": 125000.50
}
```

### Example 2: Sum Multiple Columns

**Input**:
```
Range: "A1:D10"
```

**Output**:
```json
{
  "Sum": 5000
}
```

## Business Rules

1. **Numeric Only**: Only numeric values included in sum
2. **Ignores Text**: Text values ignored
3. **Empty Cells**: Empty cells treated as zero

## Related Catalog Requests

- [Range Sum](pages/RangeSum.md) - Shared drive version
- [Columns Sum](pages/ColumnsSum.md) - Sum specific columns

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/functions/sum`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

