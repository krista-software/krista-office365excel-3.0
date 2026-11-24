# Get Add Rows to Table Result

## Overview

Retrieves the results of an asynchronous "Add Rows to Table Async" operation using the Task ID returned by that operation.

## Request Details

- **Area**: Shared
- **Type**: WAIT_FOR_EVENT
- **Retry Support**: ✅ Yes - Automatically retries on transient failures

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| eventName | Text | Yes | The name of the event (matches Task ID) | "a1b2c3d4-e5f6-7890-abcd-ef1234567890" |
| eventData | FreeForm | Yes | The data returned by the async operation | (system-provided) |
| taskID | Text | Yes | The Task ID from Add Rows to Table Async | "a1b2c3d4-e5f6-7890-abcd-ef1234567890" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sheet Data | List&lt;Composite&gt; | The rows that were added to the table |

### Output Example

```json
{
  "Sheet Data": [
    ["EMP001", "John Doe", "Engineering"],
    ["EMP002", "Jane Smith", "Marketing"]
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Task ID doesn't match event | "Invalid task ID" | Verify Task ID is correct |

## Error Handling

### Logic Errors (LOGIC_ERROR)
**Cause**: Task ID mismatch
**Resolution**: Use correct Task ID from async operation

## Usage Examples

### Example 1: Retrieve Async Results

**Input**:
```
taskID: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
```

**Output**:
```json
{
  "Sheet Data": [
    ["EMP001", "John Doe", "Engineering"],
    ["EMP002", "Jane Smith", "Marketing"]
  ]
}
```

## Business Rules

1. **Event-Based**: Waits for async operation to complete
2. **Task ID Match**: eventName must match taskID
3. **One-Time Retrieval**: Results retrieved once per task

## Limitations

1. **Requires Task ID**: Must have Task ID from async operation
2. **Event-Driven**: Waits for event to fire

## Best Practices

### 1. Store Task ID
Keep Task ID from async operation to retrieve results.

### 2. Handle Timeouts
Implement timeout handling for long-running operations.

## Common Use Cases

### 1. Check Async Completion
```
Scenario: Verify bulk import completed
Action: Retrieve results with Task ID
Result: Confirmation of rows added
```

## Related Catalog Requests

- [Add Rows to Table Async](pages/AddRowsToTableAsync.md) - Start async operation
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Synchronous version

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog
- **Type**: WAIT_FOR_EVENT

### Telemetry Metrics
- **EVENT_RECEIVED**: Event received for task
- **RESULTS_RETRIEVED**: Results successfully retrieved

## See Also

- [Add Rows to Table Async](pages/AddRowsToTableAsync.md)
- [Add Rows to Table](pages/AddRowsToTableShared.md)

