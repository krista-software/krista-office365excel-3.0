# Get Add Row With Index Result

## Overview

Retrieves the results of the asynchronous "Add Row With Index Async" operation.

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
| Sheet Data | List&lt;Composite&gt; | The rows that were inserted |

## Usage Examples

### Example 1: Retrieve Results

**Input**:
```
taskID: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
```

**Output**:
```json
{
  "Sheet Data": [
    ["EMP003", "Bob Johnson", "Sales"]
  ]
}
```

## Business Rules

1. **Event-Based**: Waits for async operation to complete
2. **Task ID Match**: eventName must match taskID

## Related Catalog Requests

- [Add Row With Index Async](pages/AddRowWithIndexAsync.md) - Start async operation

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

## See Also

- [Add Row With Index Async](pages/AddRowWithIndexAsync.md)
- [Extension Configuration](pages/ExtensionConfiguration.md)

