# Add Rows Async To Table Progress Update

## Overview

Triggers when an update on adding rows to an Excel table is available. Provides progress updates approximately every 20% of total rows to add.

## Request Details

- **Area**: Table
- **Type**: WAIT_FOR_EVENT
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| eventName | Text | Yes | Event name (matches Task ID) | "a1b2c3d4-e5f6-7890-abcd-ef1234567890" |
| eventData | FreeForm | Yes | Event data (system-provided) | (system-provided) |
| Task Id | Identifier | Yes | Task ID from batch async operation | "a1b2c3d4-e5f6-7890-abcd-ef1234567890" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Message | Paragraph | Progress status message |
| Current Index | Number | Index of last record added |
| Progress | Percentage | Completion percentage |
| Completed | Boolean | Indicates if operation is complete |

## Usage Examples

### Example 1: Progress Update

**Input**:
```
Task Id: "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
```

**Output**:
```json
{
  "Message": "Added 200 of 1000 rows",
  "Current Index": 200,
  "Progress": 20,
  "Completed": false
}
```

## Business Rules

1. **Progress Intervals**: Updates sent ~every 20%
2. **Event-Based**: Waits for progress events
3. **Multiple Updates**: Receives multiple updates until complete

## Related Catalog Requests

- [Add Rows To Table In Batch Async](pages/AddRowsToTableInBatchAsync.md) - Start batch operation
- [Add Rows Async To Table Result Available](pages/AddRowsAsyncToTableResultAvailable.md) - Final results

## Technical Implementation

### Helper Class
- **Class**: TableArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog

## See Also

- [Add Rows To Table In Batch Async](pages/AddRowsToTableInBatchAsync.md)
- [Extension Configuration](pages/ExtensionConfiguration.md)

