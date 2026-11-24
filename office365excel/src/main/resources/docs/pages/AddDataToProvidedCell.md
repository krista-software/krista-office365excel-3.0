# Add Data to Provided Cell

## Overview

Adds or updates data in a specific cell within a worksheet of an Excel file stored in a shared drive. This operation allows you to write a single value to a designated cell location.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the sheet | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | The name of the sheet containing the cell | "Sheet1" |
| Cell Id | Text | Yes | The ID of the cell where data will be added (e.g., A1, B2) | "A1" |
| Data | Text | Yes | The data to be added to the specified cell | "Hello World" |

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

#### Cell Id
- **Format**: Excel cell reference (column letter + row number)
- **Pattern**: `[A-Z]+[0-9]+` (e.g., A1, AA100, Z999)
- **Case Sensitivity**: Not case-sensitive (A1 = a1)
- **Range**: Single cell only (ranges not supported)
- **Example**: `A1`, `B5`, `AA10`, `Z999`

#### Data
- **Format**: Text value to write to the cell
- **Data Types**: Supports text, numbers, dates (as text)
- **Empty String**: Allowed (clears the cell)
- **Special Characters**: Supported
- **Formulas**: Can include Excel formulas (start with =)
- **Example**: `"Hello World"`, `"12345"`, `"=SUM(A1:A10)"`

> **📝 Note**: To write a formula, start the data with an equals sign (=). For example: `"=SUM(A1:A10)"`

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell Information | FreeForm | Object containing cell ID and the value that was written |

### Cell Information Structure

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| Cell ID | Text | The identifier of the cell | "A1" |
| Value | Text | The value that was written to the cell | "Hello World" |

### Output Example

```json
{
  "Cell Information": {
    "Cell ID": "A1",
    "Value": "Hello World"
  }
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Cell Id is empty or null | "Cell ID is required" | Provide a valid cell ID (e.g., A1) |
| Data is null | "Data is required" | Provide data value (empty string is allowed) |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and user permissions |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Sheet not found | "Sheet '{sheetName}' not found in the file" | Verify sheet name (case-sensitive) |
| Invalid cell ID format | "Invalid cell ID format" | Use valid Excel cell reference (e.g., A1, B2) |
| Cell out of range | "Cell reference is out of valid range" | Use valid cell reference within Excel limits |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters
- Invalid cell ID format (e.g., "1A" instead of "A1")
- Null data parameter (empty string is allowed)

**Error Message**: `"Drive ID is required"`, `"Cell ID is required"`, `"Invalid cell ID format"`

**Resolution**:
1. Verify all required parameters are provided
2. Check cell ID follows Excel format (letter + number)
3. Ensure data parameter is not null (use empty string to clear cell)

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Sheet doesn't exist in the specified file
- File not accessible or doesn't exist
- Drive not accessible
- Cell reference is out of valid range
- Insufficient permissions to modify file
- File is locked or in use

**Error Message**: `"Sheet '{sheetName}' not found"`, `"File is locked or in use"`

**Resolution**:
1. Verify sheet name is correct and case-sensitive
2. Ensure file exists and is accessible
3. Check user has write permissions to the file
4. Verify file is not locked or checked out
5. Ensure cell ID is within valid Excel range

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)
- Concurrent modification conflicts

**Error Message**: `"Failed to add data to cell"`, `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited
5. Ensure no concurrent modifications

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions
- User doesn't have write access to the file
- File is read-only

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite.All permission
3. Check user has write access to the file
4. Ensure file is not read-only

## Usage Examples

### Example 1: Write Text Value

**Scenario**: Write a text value to a cell

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Sheet1"
Cell Id: "A1"
Data: "Hello World"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "A1",
    "Value": "Hello World"
  }
}
```

**Result**: Text "Hello World" written to cell A1

### Example 2: Write Numeric Value

**Scenario**: Write a number to a cell

**Input**:
```
Cell Id: "B5"
Data: "1250.50"
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

**Result**: Number 1250.50 written to cell B5 (Excel may auto-format as number)

### Example 3: Write Formula

**Scenario**: Write a SUM formula to a cell

**Input**:
```
Cell Id: "C10"
Data: "=SUM(C1:C9)"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "C10",
    "Value": "=SUM(C1:C9)"
  }
}
```

**Result**: Formula written to cell C10, Excel will calculate the result

### Example 4: Clear Cell

**Scenario**: Clear the contents of a cell

**Input**:
```
Cell Id: "D5"
Data: ""
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "D5",
    "Value": ""
  }
}
```

**Result**: Cell D5 cleared (empty)

### Example 5: Write Date

**Scenario**: Write a date value to a cell

**Input**:
```
Cell Id: "E2"
Data: "2024-01-15"
```

**Output**:
```json
{
  "Cell Information": {
    "Cell ID": "E2",
    "Value": "2024-01-15"
  }
}
```

**Result**: Date written to cell E2 (Excel may auto-format as date)

## Business Rules

1. **Overwrites Existing**: Replaces any existing value in the cell
2. **Single Cell Only**: Can only write to one cell per request
3. **Auto-Formatting**: Excel may auto-format values (numbers, dates)
4. **Formula Support**: Formulas starting with = are evaluated by Excel
5. **Empty String Clears**: Empty string clears the cell contents
6. **No Formatting**: Does not apply cell formatting (colors, borders, etc.)
7. **Immediate Effect**: Changes take effect immediately
8. **Version History**: OneDrive maintains version history for recovery

## Limitations

1. **Single Cell Only**: Cannot write to multiple cells in one request
2. **No Formatting**: Cannot set cell formatting, colors, or styles
3. **Text Only**: All data is written as text (Excel auto-converts types)
4. **No Validation**: Excel data validation rules are not enforced
5. **No Protection**: Cannot write to protected cells
6. **Shared Drives Only**: This request is for shared drive files only
7. **No Batch**: Cannot batch multiple cell updates in one request

## Best Practices

### 1. Validate Cell References
Always validate cell ID format before making the request to avoid errors.

### 2. Use Formulas Wisely
When writing formulas, ensure referenced cells exist and contain appropriate data.

### 3. Handle Auto-Formatting
Be aware that Excel may auto-format values (e.g., "01/15/2024" becomes a date).

### 4. Clear Cells Properly
Use empty string ("") to clear cells, not null.

### 5. Check File Lock Status
Ensure file is not locked before attempting to write data.

### 6. Use Specific Cell References
Always use specific cell references (A1, B2) rather than named ranges.

## Common Use Cases

### 1. Update Single Value
```
Scenario: Update a status cell in a tracking sheet
Action: Write new status value to specific cell
Result: Status updated in Excel
```

### 2. Write Calculation Result
```
Scenario: Write calculated value from external system
Action: Write result to designated cell
Result: Calculation result available in Excel
```

### 3. Set Formula
```
Scenario: Add a SUM formula to total column
Action: Write formula to cell
Result: Excel calculates and displays sum
```

### 4. Clear Cell
```
Scenario: Remove outdated value from cell
Action: Write empty string to cell
Result: Cell cleared
```

## Related Catalog Requests

- [Get Cell Value](pages/GetCellValue.md) - Read cell values
- [Add Multiple Cell Data](pages/AddMultipleCellData.md) - Write to multiple cells
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add entire rows
- [Get Rows for Range](pages/GetRowsForRange.md) - Read multiple cells
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get File ID

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation and cell reference verification
- **Service**: Delegates to WorksheetService for cell operations

### Microsoft Graph API Endpoint
- **Endpoint**: `PATCH /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{cellId}')`
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **CELL_DATA_WRITTEN**: Logged when data is written
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Invalid cell ID format" Error

**Cause**: Cell ID doesn't follow Excel format

**Solution**:
1. Use format: Column letter(s) + Row number (e.g., A1, AA10)
2. Ensure no spaces or special characters
3. Verify column letters come before row number

### Issue: Data Not Formatted as Expected

**Cause**: Excel auto-formatting

**Solution**:
1. This is expected Excel behavior
2. Pre-format cells in Excel if specific formatting is needed
3. Use text format for cells that should remain as text

### Issue: Formula Not Calculating

**Cause**: Formula syntax error or referenced cells don't exist

**Solution**:
1. Verify formula syntax is correct
2. Ensure referenced cells exist
3. Check for circular references
4. Open file in Excel to see formula errors

### Issue: "File is locked" Error

**Cause**: File is open in Excel or locked by another process

**Solution**:
1. Close the file in Excel or Excel Online
2. Wait a few minutes and retry
3. Check if file is checked out in SharePoint
4. Verify no other processes are accessing the file

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get Cell Value](pages/GetCellValue.md) - Read cell values
- [Add Multiple Cell Data](pages/AddMultipleCellData.md) - Write multiple cells
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add rows

