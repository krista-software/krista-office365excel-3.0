# Get Cell Value

## Overview

Retrieves the value and metadata of a specific cell from a worksheet in an Excel file stored in the user's personal OneDrive. This operation returns both the cell identifier and its current value.

## Request Details

- **Area**: Sheet
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File ID | Text | Yes | The ID of the Excel file containing the sheet | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | The name of the sheet containing the cell | "Sheet1" |
| Cell ID | Text | Yes | The ID of the cell to retrieve (e.g., A1, B2) | "A2" |

### Parameter Details

#### File ID
- **Format**: Unique identifier for the Excel file in OneDrive
- **How to obtain**: Use "Get All Excel Files For Name" or "Get Excel File Information" catalog request
- **Validation**: Must be a valid Excel file ID accessible by the authenticated user
- **Example**: `01C8MOQGI3YRBDXAQERECWE3W4XCF`

#### Sheet Name
- **Format**: Name of the worksheet as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Sheet must exist in the Excel file
- **Special Characters**: Supported, but avoid special characters when possible
- **Example**: `Sheet1`, `Sales Data`, `Q1 Report`

#### Cell ID
- **Format**: Excel cell reference (column letter + row number)
- **Pattern**: `[A-Z]+[0-9]+` (e.g., A1, AA100, Z999)
- **Case Sensitivity**: Not case-sensitive (A1 = a1)
- **Range**: Single cell only (ranges not supported)
- **Example**: `A1`, `B5`, `AA10`, `Z999`

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell Information | FreeForm | Object containing cell ID and its value |

### Cell Information Structure

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| Cell ID | Text | The identifier of the cell | "A2" |
| Value | Text | The current value in the cell | "Hello World" |

### Output Example

```json
{
  "Cell Information": {
    "Cell ID": "A2",
    "Value": "Hello World"
  }
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| File ID is empty or null | "File ID is required" | Provide a valid File ID |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Cell ID is empty or null | "Cell ID is required" | Provide a valid cell ID (e.g., A1) |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Sheet not found | "Sheet '{sheetName}' not found in the file" | Verify sheet name (case-sensitive) |
| Invalid cell ID format | "Invalid cell ID format" | Use valid Excel cell reference (e.g., A1, B2) |
| Cell out of range | "Cell reference is out of valid range" | Use valid cell reference within Excel limits |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters (File ID, Sheet Name, Cell ID)
- Invalid cell ID format (e.g., "1A" instead of "A1")
- Empty or whitespace-only parameters

**Error Message**: `"File ID is required"`, `"Sheet Name is required"`, or `"Cell ID is required"`

**Resolution**:
1. Verify all required parameters are provided
2. Check cell ID follows Excel format (letter + number)
3. Ensure no parameters are empty or null

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Sheet doesn't exist in the specified file
- File not accessible or doesn't exist
- Cell reference is out of valid range
- Insufficient permissions to read file

**Error Message**: `"Sheet '{sheetName}' not found in the file"` or `"File not found or inaccessible"`

**Resolution**:
1. Verify sheet name is correct and case-sensitive
2. Ensure file exists and is accessible
3. Check user has read permissions to the file
4. Verify cell ID is within valid Excel range

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to retrieve cell value"` or `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions
- User doesn't have access to the file
- File is in a location user cannot access

**Error Message**: `"Unauthorized access"` or `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite permission
3. Check user has read access to the file
4. Ensure file is in user's OneDrive

## Usage Examples

### Example 1: Get Simple Text Value

**Scenario**: Retrieve a text value from a cell

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sheet1"
Cell ID: "A2"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "A2",
    "Value": "John Doe"
  }
}
```

**Result**: Successfully retrieved text value "John Doe" from cell A2

### Example 2: Get Numeric Value

**Scenario**: Retrieve a numeric value from a cell

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sales Data"
Cell ID: "B5"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "B5",
    "Value": "1250.50"
  }
}
```

**Result**: Successfully retrieved numeric value 1250.50 from cell B5

### Example 3: Get Formula Result

**Scenario**: Retrieve the calculated result of a formula

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Calculations"
Cell ID: "C10"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "C10",
    "Value": "150"
  }
}
```

**Result**: Retrieved the calculated result (150) of the formula in cell C10

> **📝 Note**: The operation returns the calculated value, not the formula itself.

### Example 4: Get Empty Cell

**Scenario**: Retrieve value from an empty cell

**Input**:
```
File ID: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sheet1"
Cell ID: "Z99"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "Z99",
    "Value": ""
  }
}
```

**Result**: Empty string returned for empty cell

## Business Rules

1. **Formula Values**: Returns the calculated result of formulas, not the formula text
2. **Data Types**: All values are returned as text strings, regardless of original type
3. **Formatting**: Cell formatting is not included in the response
4. **Empty Cells**: Empty cells return an empty string value
5. **Error Values**: Excel error values (e.g., #DIV/0!, #N/A) are returned as-is
6. **Date Values**: Dates are returned in Excel's serial number format or as formatted text
7. **Boolean Values**: TRUE/FALSE values are returned as text strings

## Limitations

1. **Single Cell Only**: Cannot retrieve multiple cells or ranges in one request
2. **No Formatting**: Cell formatting, colors, and styles are not returned
3. **No Formula Text**: Cannot retrieve the formula itself, only the calculated result
4. **Text Representation**: All values are returned as text, type conversion may be needed
5. **No Metadata**: Additional cell metadata (comments, notes) is not included
6. **Personal OneDrive Only**: This request is for personal OneDrive files (use Shared Area for shared drives)

## Best Practices

### 1. Validate Cell References
Always validate cell ID format before making the request to avoid errors.

### 2. Handle Empty Cells
Check for empty string values and handle them appropriately in your workflow.

### 3. Type Conversion
Convert string values to appropriate types (number, date, boolean) as needed in your application.

### 4. Error Value Handling
Check for Excel error values (#DIV/0!, #N/A, #REF!, etc.) and handle them gracefully.

### 5. Use Specific Sheet Names
Always specify the exact sheet name to avoid ambiguity in multi-sheet workbooks.

### 6. Cache File Metadata
Cache file and sheet information to avoid repeated lookups for the same file.

## Common Use Cases

### 1. Read Configuration Value
```
Scenario: Read a configuration value from a specific cell
Action: Get cell value from known cell location
Result: Configuration value retrieved for use in workflow
```

### 2. Validate Data Entry
```
Scenario: Check if a specific cell contains expected value
Action: Retrieve cell value and compare with expected value
Result: Validation passed or failed based on comparison
```

### 3. Read Calculation Result
```
Scenario: Get the result of a complex formula calculation
Action: Retrieve cell containing formula result
Result: Calculated value available for further processing
```

### 4. Monitor Cell Changes
```
Scenario: Periodically check if a cell value has changed
Action: Retrieve cell value and compare with previous value
Result: Detect changes and trigger appropriate actions
```

## Related Catalog Requests

- [Add Data to Cell](pages/AddDataToCell.md) - Update a cell value
- [Get Rows For Range](pages/GetRowsForRange.md) - Get multiple cells in a range
- [Find Value](pages/FindValue.md) - Search for a value across the sheet
- [Cell Data](pages/CellData.md) - Shared Area version for shared drive files
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Find File ID

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation performed before API call
- **Service**: Delegates to WorksheetService for cell operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{cellId}')`
- **Permissions Required**: Files.ReadWrite

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Sheet not found" Error

**Cause**: Sheet name doesn't match or sheet doesn't exist

**Solution**:
1. Verify sheet name is spelled correctly (case-sensitive)
2. Use "Get Sheets Information" to list available sheets
3. Check for extra spaces in sheet name

### Issue: "Invalid cell ID format" Error

**Cause**: Cell ID doesn't follow Excel format

**Solution**:
1. Use format: Column letter(s) + Row number (e.g., A1, AA10)
2. Ensure no spaces or special characters
3. Verify column letters come before row number

### Issue: Empty Value Returned Unexpectedly

**Cause**: Cell is actually empty or contains only whitespace

**Solution**:
1. Open the Excel file and verify cell contents
2. Check if cell contains formula that evaluates to empty
3. Verify you're referencing the correct cell

### Issue: Formula Instead of Value

**Cause**: This is not possible with this request (it always returns values)

**Solution**:
If you need the formula text, this request cannot provide it. Consider using Microsoft Graph API directly or alternative methods.

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Add Data to Cell](pages/AddDataToCell.md) - Update cell values
- [Get Rows For Range](pages/GetRowsForRange.md) - Get multiple cells at once

