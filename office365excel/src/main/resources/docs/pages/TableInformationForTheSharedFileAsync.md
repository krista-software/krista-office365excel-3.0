# Table Information For The Shared File Async

## Overview

Asynchronously retrieves information about all tables in an Excel file stored in a shared drive.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| fileId | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| driveId | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| taskID | Text | Task ID to retrieve results |

## Usage Examples

### Example 1: Get Table Information Async

**Input**:
```
fileId: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
driveId: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
```

**Output**:
```json
{
  "taskID": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

## Business Rules

1. **Asynchronous**: Runs in background
2. **Task ID Required**: Use Task ID to retrieve results
3. **All Tables**: Returns information for all tables in file

## Related Catalog Requests

- [Get Tables Information For Shared File](pages/GetTablesInformationForSharedFile.md) - Retrieve results

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/tables`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)

