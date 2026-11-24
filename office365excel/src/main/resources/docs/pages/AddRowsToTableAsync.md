# Add Rows to Table Async

## Overview

Asynchronously adds rows of data to an Excel table in a shared drive file. This operation runs in the background and returns a Task ID that can be used to retrieve the results later.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the table | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Table Name | Text | Yes | The name of the table to add rows to | "EmployeeTable" |
| Rows | List&lt;Text&gt; | No | List of comma-separated row data | ["John,Doe,30", "Jane,Smith,25"] |
| Rows From Csv | List&lt;File&gt; | No | CSV files containing rows to add | [data.csv] |

> **📝 Note**: At least one of "Rows" or "Rows From Csv" must be provided.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Task ID | Text | Unique identifier to retrieve the operation result |

### Output Example

```json
{
  "Task ID": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Both Rows and Rows From Csv are empty | "Row data not provided" | Provide at least one data source |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing required parameters or no row data provided
**Resolution**: Provide all required parameters and at least one data source

### Logic Errors (LOGIC_ERROR)
**Cause**: Table doesn't exist or file not accessible
**Resolution**: Verify table name and file accessibility

### System Errors (SYSTEM_ERROR)
**Cause**: Microsoft Graph API failures
**Resolution**: Retry operation

### Authorization Errors
**Cause**: Authentication or permission issues
**Resolution**: Re-authenticate, verify permissions

## Usage Examples

### Example 1: Add Rows Using Text List

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Table Name: "EmployeeTable"
Rows: ["EMP001,John Doe,Engineering", "EMP002,Jane Smith,Marketing"]
Rows From Csv: null
```

**Output**:
```json
{
  "Task ID": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

**Result**: Task ID returned immediately, operation runs in background

### Example 2: Add Rows from CSV File

**Input**:
```
Rows: null
Rows From Csv: [employees.csv]
```

**Output**:
```json
{
  "Task ID": "b2c3d4e5-f6g7-8901-bcde-fg2345678901"
}
```

**Result**: CSV file processed asynchronously

### Example 3: Retrieve Results

Use the "Get Add Rows to Table Result" catalog request with the Task ID to retrieve the results.

## Business Rules

1. **Asynchronous Execution**: Operation runs in background
2. **Task ID Required**: Must use Task ID to retrieve results
3. **No Immediate Result**: Results not available immediately
4. **Append Only**: Rows are appended to the end of the table
5. **Order Preserved**: Rows added in the order provided

## Limitations

1. **No Immediate Feedback**: Cannot get results immediately
2. **Task ID Management**: Must track Task ID to retrieve results
3. **Shared Drives Only**: For shared drive files only
4. **Table Required**: Cannot add to regular ranges

## Best Practices

### 1. Store Task ID
Save the Task ID to retrieve results later.

### 2. Use for Large Datasets
Use async operations for large datasets to avoid timeouts.

### 3. Poll for Results
Use "Get Add Rows to Table Result" to check when operation completes.

### 4. Handle Errors
Implement error handling for async operation failures.

## Common Use Cases

### 1. Bulk Data Import
```
Scenario: Import large dataset without blocking
Action: Use async operation with CSV file
Result: Task ID returned, import runs in background
```

### 2. Long-Running Operations
```
Scenario: Add thousands of rows to table
Action: Use async to avoid timeout
Result: Operation completes in background
```

## Related Catalog Requests

- [Get Add Rows to Table Result](pages/GetAddRowsToTableResult.md) - Retrieve async results
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Synchronous version
- [Get Table Rows](pages/GetTableRows.md) - View table data

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog
- **Service**: Delegates to WorkbookService for async table operations

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /drives/{driveId}/items/{fileId}/workbook/tables/{tableName}/rows/add`
- **Execution**: Asynchronous with event handler
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **ASYNC_TASK_CREATED**: Task ID created
- **ROWS_ADDED_COUNT**: Number of rows added (in background)

## Troubleshooting

### Issue: Task ID Not Working

**Solution**: Ensure you're using the correct Task ID with "Get Add Rows to Table Result"

### Issue: Results Not Available

**Solution**: Wait for async operation to complete before retrieving results

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)
- [Get Add Rows to Table Result](pages/GetAddRowsToTableResult.md)
- [Add Rows to Table](pages/AddRowsToTableShared.md)

