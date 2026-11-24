# Execute Formula

## Overview

Executes an Excel formula and returns the result without modifying the worksheet.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Formula | Text | Yes | Excel formula to execute | "=SUM(A1:A10)" |
| Formula Name | Text | Yes | Unique name for the formula | "TotalSales" |
| Comment | Text | Yes | Description of the formula | "Calculate total sales" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Formula Result | FreeForm | Result of executing the formula |

## Usage Examples

### Example 1: Execute SUM Formula

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Formula: "=SUM(A1:A10)"
Formula Name: "TotalSales"
Comment: "Calculate total sales for Q1"
```

**Output**:
```json
{
  "Formula Result": 15000
}
```

### Example 2: Execute VLOOKUP Formula

**Input**:
```
Formula: "=VLOOKUP(\"John\",A1:C10,2,FALSE)"
Formula Name: "FindEmployee"
Comment: "Find employee age"
```

**Output**:
```json
{
  "Formula Result": 30
}
```

## Business Rules

1. **No Modification**: Does not modify worksheet
2. **Formula Name Sanitized**: Invalid characters removed from name
3. **Supports All Functions**: Any valid Excel formula

## Limitations

1. **Read-Only**: Cannot modify worksheet data
2. **Single Result**: Returns single value

## Related Catalog Requests

- [Execute Formula For Shared File](pages/ExecuteFormulaForSharedFile.md) - Shared drive version

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/functions/{functionName}`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

