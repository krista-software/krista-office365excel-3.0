# Get Tables Information For Shared File

## Overview

Retrieves the results of the asynchronous "Table Information For The Shared File Async" operation.

## Request Details

- **Area**: Shared
- **Type**: WAIT_FOR_EVENT
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| eventName | Text | Yes | Event name (matches Task ID) | "a1b2c3d4-e5f6-7890-abcd-ef1234567890" |
| eventData | FreeForm | Yes | Event data (system-provided) | (system-provided) |
| taskID | Text | Yes | Task ID from async operation | "a1b2c3d4-e5f6-7890-abcd-ef1234567890" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Tables Details | List&lt;Entity(Table)&gt; | List of table entities with details |

### Table Entity Structure

| Property | Type | Description |
|----------|------|-------------|
| Table Name | Text | Name of the table |
| Table ID | Text | Unique identifier for the table |
| Sheet Name | Text | Worksheet containing the table |
| Row Count | Number | Number of rows in the table |

## Usage Examples

### Example 1: Retrieve Table Information

**Input**:
```
taskID: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
```

**Output**:
```json
{
  "Tables Details": [
    {
      "Table Name": "EmployeeTable",
      "Table ID": "table1",
      "Sheet Name": "Sheet1",
      "Row Count": 100
    },
    {
      "Table Name": "SalesTable",
      "Table ID": "table2",
      "Sheet Name": "Sheet2",
      "Row Count": 250
    }
  ]
}
```

## Business Rules

1. **Event-Based**: Waits for async operation to complete
2. **All Tables**: Returns all tables in the file
3. **Task ID Match**: eventName must match taskID

## Related Catalog Requests

- [Table Information For The Shared File Async](pages/TableInformationForTheSharedFileAsync.md) - Start async operation

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

## See Also

- [Table Information For The Shared File Async](pages/TableInformationForTheSharedFileAsync.md)
- [Extension Configuration](pages/ExtensionConfiguration.md)

