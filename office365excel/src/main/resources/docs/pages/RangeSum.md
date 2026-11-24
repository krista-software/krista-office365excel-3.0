# Range Sum

## Overview

Calculates the sum of all numeric values in a specified range within a worksheet of an Excel file stored in a shared drive. This operation is equivalent to Excel's SUM function.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the sheet | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | The name of the sheet containing the range to sum | "Sheet1" |
| Range | Text | Yes | The Excel range to sum (e.g., A1:B10) | "A1:B10" |

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
- **Format**: Excel range notation (e.g., A1:B10, C5:C20)
- **Pattern**: `[Column][Row]:[Column][Row]`
- **Single Cell**: Can be a single cell (e.g., A1)
- **Multiple Columns**: Can span multiple columns (e.g., A1:D10)
- **Validation**: Must be a valid Excel range
- **Example**: `A1:B10`, `C5:C20`, `A1:Z100`

> **📝 Note**: Range format is case-insensitive (A1:B10 = a1:b10)

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sum | Number | The sum of all numeric values in the specified range |

### Output Example

```json
{
  "Sum": 15750.50
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Range is empty or null | "Range is required" | Provide a valid Excel range |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and user permissions |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Sheet not found | "Sheet '{sheetName}' not found in the file" | Verify sheet name (case-sensitive) |
| Invalid range format | "Invalid range format" | Use valid Excel range notation (e.g., A1:B10) |
| Range out of bounds | "Range is out of valid bounds" | Ensure range is within sheet limits |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters
- Invalid range format (e.g., "1A:10B" instead of "A1:B10")
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

**Error Message**: `"Failed to calculate range sum"`, `"Microsoft Graph API error: {details}"`

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

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite.All permission
3. Check user has read access to the file

## Usage Examples

### Example 1: Sum Single Column

**Scenario**: Calculate total sales from column B

**Sheet Data**:
| A | B |
|---|---|
| Product A | 1000 |
| Product B | 2500 |
| Product C | 1750 |

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Sales"
Range: "B1:B3"
```

**Output**:
```json
{
  "Sum": 5250
}
```

**Result**: Total sales of 5250 calculated

### Example 2: Sum Multiple Columns

**Scenario**: Calculate total from multiple columns

**Sheet Data**:
| A | B | C |
|---|---|---|
| 100 | 200 | 300 |
| 150 | 250 | 350 |

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Data"
Range: "A1:C2"
```

**Output**:
```json
{
  "Sum": 1350
}
```

**Result**: Sum of all values (100+200+300+150+250+350 = 1350)

### Example 3: Sum with Mixed Data Types

**Scenario**: Sum range containing numbers and text

**Sheet Data**:
| A |
|---|
| 100 |
| Text |
| 200 |
| 300 |

**Input**:
```
Range: "A1:A4"
```

**Output**:
```json
{
  "Sum": 600
}
```

**Result**: Only numeric values are summed (100+200+300 = 600), text is ignored

### Example 4: Sum Single Cell

**Scenario**: Get value from a single cell

**Input**:
```
Range: "A1"
```

**Output**:
```json
{
  "Sum": 1000
}
```

**Result**: Single cell value returned as sum

### Example 5: Sum Empty Range

**Scenario**: Sum range with no numeric values

**Input**:
```
Range: "A1:A5"
```

**Output**:
```json
{
  "Sum": 0
}
```

**Result**: Sum is 0 when range contains no numeric values

## Business Rules

1. **Numeric Values Only**: Only numeric values are included in the sum
2. **Text Ignored**: Text values are ignored (not treated as 0)
3. **Empty Cells**: Empty cells are treated as 0
4. **Formula Results**: If cells contain formulas, their calculated values are summed
5. **Boolean Values**: TRUE = 1, FALSE = 0 in calculations
6. **Date Values**: Dates are summed as their serial number values
7. **Error Values**: Cells with errors (#DIV/0!, #N/A) are ignored
8. **Decimal Precision**: Result maintains Excel's decimal precision

## Limitations

1. **No Conditional Sum**: Cannot apply conditions (use SUMIF in Excel instead)
2. **No Multiple Ranges**: Can only sum one contiguous range per request
3. **No 3D References**: Cannot sum across multiple sheets
4. **Shared Drives Only**: This request is for shared drive files only
5. **Read-Only**: Does not modify the Excel file
6. **No Formatting**: Result is a plain number without formatting

## Best Practices

### 1. Validate Range Format
Ensure range follows Excel notation before making the request.

### 2. Handle Zero Results
Check if sum is zero and determine if it's expected (empty range vs. all zeros).

### 3. Use Specific Ranges
Specify exact ranges rather than entire columns for better performance.

### 4. Consider Data Types
Be aware that only numeric values are summed; text and errors are ignored.

### 5. Cache Results
If summing the same range repeatedly, cache results to reduce API calls.

### 6. Verify Sheet Bounds
Ensure range is within the actual data bounds of the sheet.

## Common Use Cases

### 1. Calculate Total Sales
```
Scenario: Calculate total sales from a sales column
Action: Sum the sales amount column
Result: Total sales value for reporting
```

### 2. Financial Calculations
```
Scenario: Calculate total expenses from expense sheet
Action: Sum expense range
Result: Total expenses for budget analysis
```

### 3. Inventory Totals
```
Scenario: Calculate total inventory quantity
Action: Sum quantity column
Result: Total inventory count
```

### 4. Data Validation
```
Scenario: Verify sum matches expected total
Action: Calculate sum and compare with expected value
Result: Data validation passed or failed
```

## Related Catalog Requests

- [Columns Sum](pages/ColumnsSum.md) - Sum specific columns
- [Execute Formula for Shared File](pages/ExecuteFormulaForSharedFile.md) - Execute custom formulas
- [Get Rows For Range](pages/GetRowsForRangeShared.md) - Get data from range
- [Get Cell Value](pages/GetCellValue.md) - Get single cell value
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get File ID

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation and range format verification
- **Service**: Delegates to WorksheetService for range operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{range}')`
- **Processing**: Client-side sum calculation on retrieved range data
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **RANGE_SUM_CALCULATED**: Logged when sum is calculated
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Invalid range format" Error

**Cause**: Range doesn't follow Excel notation

**Solution**:
1. Use format: Column letter(s) + Row number (e.g., A1:B10)
2. Ensure colon (:) separates start and end cells
3. Verify column letters come before row numbers

### Issue: Sum is Zero Unexpectedly

**Cause**: Range contains no numeric values or is empty

**Solution**:
1. Verify range contains numeric data
2. Check for text values that look like numbers
3. Ensure range is not empty
4. Verify range coordinates are correct

### Issue: Sum Doesn't Match Excel

**Cause**: Different handling of text, errors, or hidden rows

**Solution**:
1. Verify Excel formula matches the range
2. Check for hidden rows or filtered data
3. Ensure no error values in range
4. Verify decimal precision settings

### Issue: "Range is out of valid bounds" Error

**Cause**: Range exceeds sheet dimensions

**Solution**:
1. Verify sheet has data in the specified range
2. Check range coordinates are within Excel limits
3. Use "Get Rows For Range" to inspect sheet data first

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Columns Sum](pages/ColumnsSum.md) - Sum specific columns
- [Execute Formula for Shared File](pages/ExecuteFormulaForSharedFile.md) - Execute custom formulas

