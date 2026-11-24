# Get Complete Row Info

## Overview

Retrieves all data from a specific row in a worksheet based on the row index.

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
| Row Index | Number | Yes | Row position (1-based) | 5 |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Row Data | FreeForm | Complete row data |

## Usage Examples

### Example 1: Get Row Data

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Employees"
Row Index: 5
```

**Output**:
```json
{
  "Row Data": ["EMP001", "John Doe", "Engineering", "75000"]
}
```

## Business Rules

1. **1-Based Index**: Row Index starts at 1
2. **All Columns**: Returns all column values in the row
3. **Empty Cells**: Empty cells returned as empty strings

## Limitations

1. **Single Row**: Retrieves one row at a time
2. **Shared Drives Only**: For shared drive files only

## Related Catalog Requests

- [Update Row With Index](pages/UpdateRowWithIndex.md)
- [Delete Row With Index](pages/DeleteRowWithIndex.md)
- [Get Rows for Range](pages/GetRowsForRange.md)

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='A{rowIndex}:{rowIndex}')`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

