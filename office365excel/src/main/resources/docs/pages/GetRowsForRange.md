# Get Rows for Range

## Overview

Retrieves all rows of data from a specified range within a worksheet of an Excel file stored in a shared drive. This
operation returns the complete data set from the range, including all cell values.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description                                          | Example                              |
|----------------|------|----------|------------------------------------------------------|--------------------------------------|
| Drive Id       | Text | Yes      | The ID of the drive where the Excel file is located  | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"   |
| File Id        | Text | Yes      | The ID of the Excel file containing the sheet        | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name     | Text | Yes      | The name of the sheet containing the range           | "Sheet1"                             |
| Range          | Text | Yes      | The Excel range to retrieve rows from (e.g., A1:D10) | "A1:D10"                             |

### Parameter Details

#### Drive Id

- **Format**: Unique identifier for the shared drive
- **How to obtain**: Use "Get All Shared Drives Information" catalog request
- **Validation**: Must be a valid drive ID accessible by the authenticated user
- **Example**: `b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8`

#### File Id

- **Format**: Unique identifier for the Excel file
- **How to obtain**: Use "Get Shared Excel Files" catalog request
- **Validation**: Must be a valid Excel file ID in the specified drive
- **Example**: `01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR`

#### Sheet Name

- **Format**: Name of the worksheet as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Sheet must exist in the Excel file
- **Example**: `Sheet1`, `Sales Data`, `Q1 Report`

#### Range

- **Format**: Excel range notation (e.g., A1:D10, C5:F20)
- **Pattern**: `[Column][Row]:[Column][Row]`
- **Single Cell**: Can be a single cell (e.g., A1)
- **Multiple Columns**: Can span multiple columns and rows
- **Validation**: Must be a valid Excel range
- **Example**: `A1:D10`, `C5:F20`, `A1:Z100`

> **📝 Note**: Range format is case-insensitive (A1:D10 = a1:d10)

## Output Parameters

| Parameter Name | Type                  | Description                                              |
|----------------|-----------------------|----------------------------------------------------------|
| Rows           | List&lt;Composite&gt; | Array of rows, where each row is an array of cell values |

### Output Structure

The output is a two-dimensional array where:

- Each element in the outer array represents a row
- Each element in the inner array represents a cell value in that row

### Output Example

```json
{
  "Rows": [
    [
      "Name",
      "Department",
      "Salary",
      "Location"
    ],
    [
      "John Doe",
      "Engineering",
      "75000",
      "New York"
    ],
    [
      "Jane Smith",
      "Marketing",
      "65000",
      "San Francisco"
    ],
    [
      "Bob Johnson",
      "Sales",
      "70000",
      "Chicago"
    ]
  ]
}
```

## Validation Rules

| Validation                  | Error Message                               | Resolution                                    |
|-----------------------------|---------------------------------------------|-----------------------------------------------|
| Drive Id is empty or null   | "Drive ID is required"                      | Provide a valid Drive ID                      |
| File Id is empty or null    | "File ID is required"                       | Provide a valid File ID                       |
| Sheet Name is empty or null | "Sheet Name is required"                    | Provide a valid sheet name                    |
| Range is empty or null      | "Range is required"                         | Provide a valid Excel range                   |
| Drive not found             | "Drive not found or inaccessible"           | Verify Drive ID and user permissions          |
| File not found              | "File not found or inaccessible"            | Verify File ID and user permissions           |
| Sheet not found             | "Sheet '{sheetName}' not found in the file" | Verify sheet name (case-sensitive)            |
| Invalid range format        | "Invalid range format"                      | Use valid Excel range notation (e.g., A1:D10) |
| Range out of bounds         | "Range is out of valid bounds"              | Ensure range is within sheet limits           |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:

- Missing required parameters
- Invalid range format (e.g., "1A:10D" instead of "A1:D10")
- Empty or whitespace-only parameters

**Error Message**: `"Drive ID is required"`, `"Range is required"`, `"Invalid range format"`

**Resolution**:

1. Verify all required parameters are provided
2. Check range follows Excel notation (Column + Row)
3. Ensure no parameters are empty or null

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:

- Sheet doesn't exist in the specified file
- File not accessible or doesn't exist
- Drive not accessible
- Range is out of sheet bounds
- Insufficient permissions to read file

**Error Message**: `"Sheet '{sheetName}' not found"`, `"Range is out of valid bounds"`

**Resolution**:

1. Verify sheet name is correct and case-sensitive
2. Ensure file exists and is accessible
3. Check range is within sheet dimensions
4. Verify user has read permissions to the file

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:

- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)
- Large range causing timeout

**Error Message**: `"Failed to retrieve rows for range"`, `"Microsoft Graph API error: {details}"`

**Resolution**:

1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited
5. Reduce range size if timeout occurs

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:

- Access token expired
- Insufficient permissions
- User doesn't have access to the file

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:

1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite.All permission
3. Check user has read access to the file

## Usage Examples

### Example 1: Get Data Table

**Scenario**: Retrieve employee data from a range

**Sheet Data** (Range A1:D4):
| Name | Department | Salary | Location |
|------|------------|--------|----------|
| John Doe | Engineering | 75000 | New York |
| Jane Smith | Marketing | 65000 | San Francisco |
| Bob Johnson | Sales | 70000 | Chicago |

**Input**:

```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Employees"
Range: "A1:D4"
```

**Output**:

```json
{
  "Rows": [
    [
      "Name",
      "Department",
      "Salary",
      "Location"
    ],
    [
      "John Doe",
      "Engineering",
      "75000",
      "New York"
    ],
    [
      "Jane Smith",
      "Marketing",
      "65000",
      "San Francisco"
    ],
    [
      "Bob Johnson",
      "Sales",
      "70000",
      "Chicago"
    ]
  ]
}
```

**Result**: Complete data table retrieved including headers

### Example 2: Get Single Column

**Scenario**: Retrieve only the names column

**Input**:

```
Range: "A1:A4"
```

**Output**:

```json
{
  "Rows": [
    [
      "Name"
    ],
    [
      "John Doe"
    ],
    [
      "Jane Smith"
    ],
    [
      "Bob Johnson"
    ]
  ]
}
```

**Result**: Single column data retrieved

### Example 3: Get Single Row

**Scenario**: Retrieve header row only

**Input**:

```
Range: "A1:D1"
```

**Output**:

```json
{
  "Rows": [
    [
      "Name",
      "Department",
      "Salary",
      "Location"
    ]
  ]
}
```

**Result**: Single row retrieved

### Example 4: Get Single Cell

**Scenario**: Retrieve a single cell value

**Input**:

```
Range: "B2"
```

**Output**:

```json
{
  "Rows": [
    [
      "Engineering"
    ]
  ]
}
```

**Result**: Single cell value returned as a row with one element

### Example 5: Get Range with Empty Cells

**Scenario**: Retrieve range containing empty cells

**Input**:

```
Range: "A1:C3"
```

**Output**:

```json
{
  "Rows": [
    [
      "Name",
      "Age",
      ""
    ],
    [
      "John",
      "30",
      "Engineer"
    ],
    [
      "Jane",
      "",
      "Manager"
    ]
  ]
}
```

**Result**: Empty cells returned as empty strings

## Business Rules

1. **Includes All Data**: Returns all cell values in the specified range
2. **Preserves Structure**: Maintains row and column structure
3. **Empty Cells**: Empty cells are returned as empty strings
4. **Formula Values**: Returns calculated values, not formulas
5. **Data Types**: All values are returned as text strings
6. **No Formatting**: Cell formatting is not included
7. **Header Rows**: Headers are included if they're in the range
8. **Order Preserved**: Rows and columns are returned in their original order

## File Size Limits

| Size Range      | Status        | Notes                                                                                                                |
|-----------------|---------------|----------------------------------------------------------------------------------------------------------------------|
| Up to 100 MB    | Supported     | Fully supported; operations complete reliably                                                                        |
| 100 MB - 150 MB | Best Effort   | May work but success is not guaranteed; large files in this range are prone to `MaxRequestDurationExceeded` timeouts |
| Above 150 MB    | Not Supported | Operations are expected to fail due to Microsoft Graph API timeout limits                                            |

> **Important**: For files above 100 MB, consider splitting the workbook into smaller files or reducing the data volume.
> Session management (see below) helps with repeated reads on the same file, but cannot overcome the inherent size limits
> of the Microsoft Graph API.

## Session Management

Starting with version 3.0.1, this request uses **workbook session caching** to improve performance on large Excel files.

### How It Works

1. **First request** to a workbook creates a Microsoft Graph workbook session via `POST /workbook/createSession`. This "
   warms up" the workbook on the server side.
2. **Subsequent requests** to the same file (within 5 minutes) reuse the cached session via the `workbook-session-id`
   header, avoiding repeated cold workbook loads.
3. If a cached session is rejected by the server (e.g., session expired or evicted), the system automatically *
   *invalidates the cache and retries** the request without a session.

### Benefits

- Eliminates repeated cold workbook loads that previously caused `MaxRequestDurationExceeded` timeouts (often surfaced
  as "Sheet not found" errors) on large Excel files
- After the first request warms the session, follow-up reads on the same file complete significantly faster
- Fully backward-compatible: if session creation fails, the request proceeds without a session (same behavior as before)

### Session Cache Details

| Property     | Value                                             |
|--------------|---------------------------------------------------|
| Cache TTL    | 5 minutes                                         |
| Cache Key    | `{driveId}\|{fileId}\|{persistChanges}`           |
| Session Type | Read-only (for query operations)                  |
| Fallback     | Automatic retry without session on session errors |

## Limitations

1. **File Size**: Maximum supported file size is 100 MB. Files between 100-150 MB may work but are not guaranteed.
   See [File Size Limits](#file-size-limits) above.
2. **Large Ranges**: Very large ranges may cause timeouts (recommended: <10,000 cells)
3. **No Formatting**: Cell formatting, colors, and styles are not returned
4. **No Formulas**: Cannot retrieve formula text, only calculated values
5. **Text Representation**: All values are returned as text
6. **No Metadata**: Additional cell metadata (comments, notes) is not included
7. **Shared Drives Only**: This request is for shared drive files only
8. **Read-Only**: Does not modify the Excel file

## Best Practices

### 1. Specify Exact Ranges

Use specific ranges rather than entire columns for better performance.

### 2. Handle Empty Cells

Check for empty string values and handle them appropriately in your workflow.

### 3. Type Conversion

Convert string values to appropriate types (number, date, boolean) as needed.

### 4. Limit Range Size

For large datasets, retrieve data in smaller chunks to avoid timeouts.

### 5. Include Headers

Include header rows in your range for self-documenting data.

### 6. Cache Results

If accessing the same range repeatedly, cache results to reduce API calls.

## Common Use Cases

### 1. Data Export

```
Scenario: Export data from Excel to external system
Action: Retrieve all data from a range
Result: Data available for processing or storage
```

### 2. Report Generation

```
Scenario: Generate report from Excel data
Action: Get data range and format for report
Result: Report created with Excel data
```

### 3. Data Validation

```
Scenario: Validate data in Excel against business rules
Action: Retrieve data and apply validation logic
Result: Data validated and errors identified
```

### 4. Data Analysis

```
Scenario: Analyze sales data from Excel
Action: Retrieve sales data range
Result: Data available for analysis and calculations
```

## Related Catalog Requests

- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add data to sheets
- [Get Cell Value](pages/GetCellValue.md) - Get single cell value
- [Range Sum](pages/RangeSum.md) - Calculate sum of range
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get File ID
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Get Drive ID

## Technical Implementation

### Helper Class

- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation and range format verification
- **Service**: Delegates to WorksheetService for range operations

### Microsoft Graph API Endpoint

- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')`
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics

- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **ROWS_RETRIEVED_COUNT**: Number of rows retrieved
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Invalid range format" Error

**Cause**: Range doesn't follow Excel notation

**Solution**:

1. Use format: Column letter(s) + Row number (e.g., A1:D10)
2. Ensure colon (:) separates start and end cells
3. Verify column letters come before row numbers

### Issue: Empty Rows Returned

**Cause**: Range contains no data or is outside data bounds

**Solution**:

1. Verify range contains data
2. Check range coordinates are correct
3. Ensure range is within sheet bounds

### Issue: "Range is out of valid bounds" Error

**Cause**: Range exceeds sheet dimensions

**Solution**:

1. Verify sheet has data in the specified range
2. Check range coordinates are within Excel limits
3. Inspect sheet to determine actual data bounds

### Issue: Timeout on Large Range

**Cause**: Range is too large

**Solution**:

1. Reduce range size
2. Retrieve data in smaller chunks
3. Consider using pagination strategy

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add data to sheets
- [Get Cell Value](pages/GetCellValue.md) - Get single cell
- [Range Sum](pages/RangeSum.md) - Calculate sum

