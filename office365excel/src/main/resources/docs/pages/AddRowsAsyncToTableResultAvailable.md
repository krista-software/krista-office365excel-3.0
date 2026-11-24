# Add Rows Async To Table Result Available

## Overview

Retrieves the final results when the asynchronous batch add rows operation completes.

## Request Details

- **Area**: Table
- **Type**: WAIT_FOR_EVENT
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| eventName | Text | Yes | Event name (Task ID + "_completion") | "a1b2c3d4-e5f6-7890-abcd-ef1234567890_completion" |
| eventData | FreeForm | Yes | Event data (system-provided) | (system-provided) |
| Task Id | Identifier | Yes | Task ID from batch async operation | "a1b2c3d4-e5f6-7890-abcd-ef1234567890" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Total Rows | Number | Total rows in table after operation |
| Rows Added | Number | Number of rows added |
| Success | Boolean | Indicates if operation succeeded |
| Message | Paragraph | Status message |
| Start Time | Time | Operation start time |
| End Time | Time | Operation end time |
| Duration | Number | Duration in milliseconds |
| Duration Text | Text | Human-readable duration |

## Usage Examples

### Example 1: Final Results

**Input**:
```
Task Id: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
```

**Output**:
```json
{
  "Total Rows": 1500,
  "Rows Added": 1000,
  "Success": true,
  "Message": "Successfully added 1000 rows",
  "Start Time": "2025-10-03T10:00:00Z",
  "End Time": "2025-10-03T10:05:30Z",
  "Duration": 330000,
  "Duration Text": "5 minutes 30 seconds"
}
```

## Business Rules

1. **Completion Event**: Triggered when operation completes
2. **Performance Metrics**: Includes timing information
3. **Success Indicator**: Shows if operation succeeded

## Related Catalog Requests

- [Add Rows To Table In Batch Async](pages/AddRowsToTableInBatchAsync.md) - Start batch operation
- [Add Rows Async To Table Progress Update](pages/AddRowsAsyncToTableProgressUpdate.md) - Track progress

## Technical Implementation

### Helper Class
- **Class**: TableArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

## See Also

- [Add Rows To Table In Batch Async](pages/AddRowsToTableInBatchAsync.md)
- [Extension Configuration](pages/ExtensionConfiguration.md)

