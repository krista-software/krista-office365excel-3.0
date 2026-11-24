# Execute Formula For Shared File

## Overview

Executes an Excel formula in a file stored in a shared drive and returns the calculated result. This operation allows you to perform calculations without modifying the file.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file where the formula will be executed | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Formula | Text | Yes | The Excel formula to execute | "=SUM(A1:A10)" |
| Formula Name | Text | Yes | A name to identify this formula | "Total Sales" |
| Comment | Text | Yes | Optional comment about the formula | "Calculate total sales for Q1" |

### Parameter Details

#### Formula
- **Format**: Excel formula starting with `=`
- **Syntax**: Standard Excel formula syntax
- **Functions**: Supports all Excel functions (SUM, AVERAGE, VLOOKUP, IF, etc.)
- **References**: Can reference cells, ranges, and named ranges
- **Example**: `=SUM(A1:A10)`, `=AVERAGE(B1:B20)`, `=IF(C1>100,"High","Low")`

> **📝 Note**: Formula must start with `=` sign

#### Formula Name
- **Purpose**: Identifier for the formula execution
- **Format**: Descriptive text
- **Example**: `"Total Revenue"`, `"Average Score"`, `"Lookup Result"`

#### Comment
- **Purpose**: Additional context about the formula
- **Format**: Free text
- **Optional**: Can be empty string
- **Example**: `"Monthly sales calculation"`, `"Q1 2024 analysis"`

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Formula Result | FreeForm | The calculated result of the formula execution |

### Output Example

**Numeric Result**:
```json
{
  "Formula Result": 12500.50
}
```

**Text Result**:
```json
{
  "Formula Result": "High"
}
```

**Boolean Result**:
```json
{
  "Formula Result": true
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive Id is required" | Provide a valid Drive Id |
| File Id is empty or null | "File Id is required" | Provide a valid File Id |
| Formula is empty or null | "Formula is required" | Provide a valid formula |
| Formula Name is empty or null | "Formula Name is required" | Provide a formula name |
| Comment is null | "Comment is required" | Provide comment (can be empty string) |
| Formula doesn't start with = | "Formula must start with =" | Add = at the beginning |
| Invalid formula syntax | "Invalid formula syntax" | Check formula syntax |
| Drive not found | "Drive not found or inaccessible" | Verify Drive Id and permissions |
| File not found | "File not found or inaccessible" | Verify File Id and permissions |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing or invalid parameters
**Resolution**: Provide all required parameters with valid formats

### Logic Errors (LOGIC_ERROR)
**Cause**: Invalid formula syntax or file not accessible
**Resolution**: Verify formula syntax and file accessibility

### System Errors (SYSTEM_ERROR)
**Cause**: Microsoft Graph API failures
**Resolution**: Retry operation

### Authorization Errors
**Cause**: Authentication or permission issues
**Resolution**: Re-authenticate, verify permissions

## Usage Examples

### Example 1: Calculate Sum

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Formula: "=SUM(A1:A10)"
Formula Name: "Total Sales"
Comment: "Sum of all sales values"
```

**Output**:
```json
{
  "Formula Result": 12500.50
}
```

**Result**: Sum of cells A1 through A10 calculated

### Example 2: Calculate Average

**Input**:
```
Formula: "=AVERAGE(B1:B20)"
Formula Name: "Average Score"
Comment: "Average of test scores"
```

**Output**:
```json
{
  "Formula Result": 85.5
}
```

**Result**: Average of cells B1 through B20 calculated

### Example 3: Conditional Formula

**Input**:
```
Formula: "=IF(C1>100,\"High\",\"Low\")"
Formula Name: "Category"
Comment: "Categorize based on value"
```

**Output**:
```json
{
  "Formula Result": "High"
}
```

**Result**: Conditional result based on cell C1 value

### Example 4: VLOOKUP Formula

**Input**:
```
Formula: "=VLOOKUP(\"EMP001\",A1:D100,2,FALSE)"
Formula Name: "Employee Name Lookup"
Comment: "Find employee name by ID"
```

**Output**:
```json
{
  "Formula Result": "John Doe"
}
```

**Result**: VLOOKUP result returned

### Example 5: Complex Calculation

**Input**:
```
Formula: "=SUM(A1:A10)*0.15"
Formula Name: "Tax Calculation"
Comment: "Calculate 15% tax on total"
```

**Output**:
```json
{
  "Formula Result": 1875.075
}
```

**Result**: Complex calculation performed

## Business Rules

1. **Read-Only**: Formula execution does not modify the file
2. **Temporary**: Result is calculated but not stored in the file
3. **Formula Syntax**: Must follow Excel formula syntax
4. **Function Support**: Supports all standard Excel functions
5. **Cell References**: Can reference any cells in the workbook
6. **No Side Effects**: Does not change any cell values

## Limitations

1. **No File Modification**: Cannot write results back to file
2. **Single Formula**: Executes one formula at a time
3. **No Array Formulas**: Does not support array formulas
4. **Shared Drives Only**: For shared drive files only
5. **Execution Time**: Complex formulas may take longer to execute
6. **No Volatile Functions**: Some volatile functions may not work as expected

## Best Practices

### 1. Validate Formula Syntax
Test formula in Excel before executing via API.

### 2. Use Descriptive Names
Provide meaningful Formula Name for tracking and debugging.

### 3. Add Comments
Include helpful comments to document formula purpose.

### 4. Handle Errors
Implement error handling for invalid formulas or calculation errors.

### 5. Optimize Complex Formulas
Break down complex formulas into simpler ones if possible.

## Common Use Cases

### 1. Calculate Totals
```
Scenario: Calculate total revenue without modifying file
Action: Execute SUM formula
Result: Total calculated and returned
```

### 2. Perform Lookups
```
Scenario: Look up value without adding formula to file
Action: Execute VLOOKUP formula
Result: Lookup result returned
```

### 3. Conditional Analysis
```
Scenario: Evaluate conditions and return result
Action: Execute IF formula
Result: Conditional result returned
```

### 4. Statistical Calculations
```
Scenario: Calculate average, median, or other statistics
Action: Execute statistical formula
Result: Statistical result returned
```

## Related Catalog Requests

- [Range Sum](pages/RangeSum.md) - Calculate sum of range
- [Columns Sum](pages/ColumnsSum.md) - Sum table columns
- [Table Vlookup](pages/TableVlookup.md) - Perform VLOOKUP
- [Get Cell Value](pages/GetCellValue.md) - Read cell values

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation, formula syntax check
- **Service**: Delegates to WorkbookService for formula execution

### Microsoft Graph API Endpoint
- **Endpoint**: `POST /drives/{driveId}/items/{fileId}/workbook/functions/{functionName}`
- **Alternative**: Uses workbook calculation API
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **FORMULA_EXECUTED**: Formula that was executed
- **EXECUTION_TIME**: Time taken to execute formula

## Troubleshooting

### Issue: "Formula must start with =" Error

**Cause**: Formula doesn't begin with equals sign

**Solution**:
Add `=` at the beginning of the formula

### Issue: "Invalid formula syntax" Error

**Cause**: Formula contains syntax errors

**Solution**:
1. Test formula in Excel first
2. Check for typos in function names
3. Verify cell references are valid
4. Ensure parentheses are balanced

### Issue: #REF! or #NAME? Error in Result

**Cause**: Formula references invalid cells or functions

**Solution**:
1. Verify cell references exist in the file
2. Check function names are spelled correctly
3. Ensure referenced ranges are valid

### Issue: Unexpected Result

**Cause**: Formula logic or cell references incorrect

**Solution**:
1. Test formula in Excel to verify logic
2. Check cell references point to correct data
3. Verify data types in referenced cells

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Range Sum](pages/RangeSum.md) - Sum range values
- [Columns Sum](pages/ColumnsSum.md) - Sum column values
- [Table Vlookup](pages/TableVlookup.md) - Perform VLOOKUP

