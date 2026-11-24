# Get Add Rows Result (Table)

## Overview

Retrieves the results of the asynchronous "Add Rows Async" operation for tables.

## Request Details

- **Area**: Table
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
| Total Rows | List&lt;Composite&gt; | The rows that were added |

## Usage Examples

### Example 1: Retrieve Results

**Input**:
```
taskID: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
```

**Output**:
```json
{
  "Total Rows": [
    ["EMP003", "Bob Johnson", "Sales"]
  ]
}
```

## Business Rules

1. **Event-Based**: Waits for async operation to complete
2. **Task ID Match**: eventName must match taskID

## Related Catalog Requests

- [Add Rows Async](pages/AddRowsAsync.md) - Start async operation

## Technical Implementation

### Helper Class
- **Class**: TableArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

## See Also

- [Add Rows Async](pages/AddRowsAsync.md)
- [Extension Configuration](pages/ExtensionConfiguration.md)

