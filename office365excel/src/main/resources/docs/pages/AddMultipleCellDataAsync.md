# Add Multiple Cell Data Async

## Overview

Asynchronously updates multiple cells in a worksheet with different values.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Cell ID Data | List&lt;{Cell Id: Text, Data: Text}&gt; | Yes | Cell IDs and values | [{"Cell Id": "A1", "Data": "Hello"}] |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Task ID | Text | Task ID to retrieve results |

## Usage Examples

### Example 1: Update Multiple Cells Async

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Sheet1"
Cell ID Data: [
  {"Cell Id": "A1", "Data": "Name"},
  {"Cell Id": "B1", "Data": "Age"}
]
```

**Output**:
```json
{
  "Task ID": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

## Business Rules

1. **Asynchronous**: Runs in background
2. **Batch Update**: Updates multiple cells in one operation
3. **Task ID Required**: Use Task ID to retrieve results

## Related Catalog Requests

- [Get Add Multiple Cell Data Result](pages/GetAddMultipleCellDataResult.md) - Retrieve results
- [Add Multiple Cell Data](pages/AddMultipleCellData.md) - Synchronous version

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

### Microsoft Graph API Endpoint
- **Endpoint**: `PATCH /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

