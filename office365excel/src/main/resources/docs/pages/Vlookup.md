# Vlookup

## Overview

Performs a VLOOKUP operation to find a value in the first column of a range and return a value from a specified column in the same row.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | File ID | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Lookup Value | Text | Yes | Value to search for | "EMP001" |
| Sheet Name | Text | Yes | Sheet name | "Employees" |
| Range | Text | Yes | Cell range to search | "A1:D100" |
| Column Index | Number | Yes | Column to return value from (1-based) | 3 |
| Exact Match | Boolean | Yes | true = exact match, false = approximate | true |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Lookup Result | FreeForm | Contains lookup value and result |

## Usage Examples

### Example 1: Find Employee Name

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Lookup Value: "EMP001"
Sheet Name: "Employees"
Range: "A1:D100"
Column Index: 2
Exact Match: true
```

**Output**:
```json
{
  "Lookup Result": {
    "Lookup Value": "EMP001",
    "Result": "John Doe"
  }
}
```

### Example 2: Find Salary

**Input**:
```
Lookup Value: "EMP001"
Column Index: 4
```

**Output**:
```json
{
  "Lookup Result": {
    "Lookup Value": "EMP001",
    "Result": "75000"
  }
}
```

## Business Rules

1. **First Column**: Searches first column of range
2. **1-Based Index**: Column Index starts at 1
3. **Exact Match**: When true, requires exact match
4. **Approximate Match**: When false, finds closest match

## Limitations

1. **Single Result**: Returns first match only
2. **Left to Right**: Can only return columns to the right

## Related Catalog Requests

- [Sheet Vlookup](pages/SheetVlookup.md) - Shared drive version
- [Table Vlookup](pages/TableVlookup.md) - Table version

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /me/drive/items/{fileId}/workbook/functions/vlookup`
- **Permissions Required**: Files.ReadWrite

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

