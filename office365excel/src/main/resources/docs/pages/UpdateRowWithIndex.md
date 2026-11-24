# Update Row With Index

## Overview

Updates an existing row at a specific position in a worksheet of an Excel file stored in a shared drive.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Row | Text | Yes | Comma-separated row data | "John,Doe,30,Engineer" |
| Row Index | Number | Yes | Row position (1-based) | 5 |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sheet Data | List&lt;Composite&gt; | The updated row data |

## Usage Examples

### Example 1: Update Row

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Employees"
Row: "EMP001,John Doe,Engineering,80000"
Row Index: 5
```

**Output**:
```json
{
  "Sheet Data": [
    ["EMP001", "John Doe", "Engineering", "80000"]
  ]
}
```

## Business Rules

1. **Overwrites Existing**: Replaces existing row data
2. **1-Based Index**: Row Index starts at 1
3. **Complete Row**: Updates entire row

## Limitations

1. **Single Row**: Updates one row at a time
2. **Shared Drives Only**: For shared drive files only

## Related Catalog Requests

- [Add Row With Index](pages/AddRowWithIndexShared.md)
- [Delete Row With Index](pages/DeleteRowWithIndex.md)
- [Get Complete Row Info](pages/GetCompleteRowInfo.md)

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `PATCH /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

