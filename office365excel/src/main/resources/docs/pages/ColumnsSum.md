# Columns Sum

## Overview

Calculates the sum of values in one or more columns of an Excel table stored in a shared drive file. This operation performs numerical summation on specified table columns.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the table | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Table Name | Text | Yes | The name of the table containing the columns to sum | "SalesTable" |
| Columns | List&lt;Text&gt; | Yes | List of column names to calculate sum for | ["Revenue", "Profit"] |

### Parameter Details

#### Table Name
- **Format**: Name of the table as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Table must exist in the Excel file
- **Example**: `Table1`, `SalesData`, `FinancialRecords`

#### Columns
- **Format**: List of column names from the table
- **Case Sensitivity**: Case-sensitive
- **Multiple Columns**: Can sum multiple columns in one request
- **Validation**: Column names must exist in the table
- **Data Type**: Columns should contain numeric values
- **Example**: `["Revenue"]`, `["Q1 Sales", "Q2 Sales", "Q3 Sales"]`

> **📝 Note**: Non-numeric values in columns are treated as 0 in the sum calculation.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Sum | Number | The sum of all numeric values in the specified columns |

### Output Example

**Single Column**:
```json
{
  "Sum": 125000.50
}
```

**Multiple Columns**:
```json
{
  "Sum": 450000.75
}
```

> **📝 Note**: When multiple columns are specified, the output is the total sum across all columns.

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Columns is empty or null | "Columns is required" | Provide at least one column name |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and permissions |
| File not found | "File not found or inaccessible" | Verify File ID and permissions |
| Table not found | "Table '{tableName}' not found" | Verify table name (case-sensitive) |
| Column not found | "Column '{columnName}' not found in table" | Verify column names (case-sensitive) |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters
- Empty Columns list
- Null values

**Error Message**: `"Drive ID is required"`, `"Columns is required"`

**Resolution**:
1. Verify all required parameters are provided
2. Ensure Columns list has at least one column name
3. Check no parameters are null

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Table doesn't exist in the file
- Column names don't exist in the table
- File not accessible
- Insufficient permissions

**Error Message**: `"Table '{tableName}' not found"`, `"Column '{columnName}' not found in table"`

**Resolution**:
1. Verify table name is correct and case-sensitive
2. Check column names are correct and case-sensitive
3. Ensure file exists and is accessible
4. Verify user has read permissions

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to calculate column sum"`, `"Microsoft Graph API error: {details}"`

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

**Table Data**:
| Product | Revenue | Profit |
|---------|---------|--------|
| Product A | 10000 | 2000 |
| Product B | 15000 | 3000 |
| Product C | 20000 | 4000 |

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Table Name: "SalesTable"
Columns: ["Revenue"]
```

**Output**:
```json
{
  "Sum": 45000
}
```

**Result**: Total revenue = 10000 + 15000 + 20000 = 45000

### Example 2: Sum Multiple Columns

**Input**:
```
Columns: ["Revenue", "Profit"]
```

**Output**:
```json
{
  "Sum": 54000
}
```

**Result**: Total = (10000 + 15000 + 20000) + (2000 + 3000 + 4000) = 54000

### Example 3: Sum Quarterly Sales

**Table Data**:
| Region | Q1 | Q2 | Q3 | Q4 |
|--------|-----|-----|-----|-----|
| North | 5000 | 6000 | 7000 | 8000 |
| South | 4000 | 5000 | 6000 | 7000 |
| East | 3000 | 4000 | 5000 | 6000 |

**Input**:
```
Columns: ["Q1", "Q2", "Q3", "Q4"]
```

**Output**:
```json
{
  "Sum": 66000
}
```

**Result**: Total sales across all quarters and regions = 66000

### Example 4: Column with Non-Numeric Values

**Table Data**:
| Product | Sales |
|---------|-------|
| Product A | 1000 |
| Product B | N/A |
| Product C | 2000 |

**Input**:
```
Columns: ["Sales"]
```

**Output**:
```json
{
  "Sum": 3000
}
```

**Result**: Non-numeric "N/A" treated as 0, sum = 1000 + 0 + 2000 = 3000

### Example 5: Empty Table

**Input**:
```
Columns: ["Revenue"]
```

**Output**:
```json
{
  "Sum": 0
}
```

**Result**: No data rows, sum = 0

## Business Rules

1. **Numeric Values Only**: Only numeric values are summed
2. **Non-Numeric Handling**: Non-numeric values treated as 0
3. **Empty Cells**: Empty cells treated as 0
4. **Multiple Columns**: Sum is calculated across all specified columns
5. **All Rows**: Includes all rows in the table
6. **Decimal Precision**: Maintains decimal precision in calculations
7. **Negative Numbers**: Negative numbers are included in sum

## Limitations

1. **Table Required**: Cannot sum columns in regular ranges (must be Excel table)
2. **No Filtering**: Cannot filter rows before summing
3. **No Grouping**: Cannot group by categories
4. **Shared Drives Only**: For shared drive files only
5. **No Conditional Sum**: Cannot apply conditions (use SUMIF in Excel)
6. **All Rows Included**: Cannot exclude specific rows

## Best Practices

### 1. Verify Column Names
Ensure column names match exactly (case-sensitive) before making request.

### 2. Use Numeric Columns
Ensure columns contain primarily numeric data for accurate results.

### 3. Handle Non-Numeric Data
Be aware that non-numeric values are treated as 0.

### 4. Validate Table Structure
Verify table exists and contains expected columns.

### 5. Check for Empty Tables
Handle cases where table has no data rows (sum = 0).

## Common Use Cases

### 1. Calculate Total Revenue
```
Scenario: Calculate total revenue from sales table
Action: Sum the Revenue column
Result: Total revenue calculated
```

### 2. Calculate Quarterly Totals
```
Scenario: Sum all quarterly sales columns
Action: Sum Q1, Q2, Q3, Q4 columns
Result: Annual total calculated
```

### 3. Financial Reporting
```
Scenario: Calculate total expenses across categories
Action: Sum multiple expense columns
Result: Total expenses calculated
```

### 4. Inventory Valuation
```
Scenario: Calculate total inventory value
Action: Sum inventory value column
Result: Total inventory value calculated
```

## Related Catalog Requests

- [Range Sum](pages/RangeSum.md) - Sum values in a range
- [Get Table Rows](pages/GetTableRows.md) - Retrieve all table data
- [Get Column Values For Keys](pages/GetColumnValuesForKeys.md) - Get specific column values
- [Table Vlookup](pages/TableVlookup.md) - Lookup values in table

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation, column existence check
- **Service**: Delegates to WorkbookService for table operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/tables/{tableName}/columns/{columnName}/dataBodyRange`
- **Processing**: Client-side summation of retrieved column values
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **COLUMNS_SUMMED_COUNT**: Number of columns summed
- **SUM_RESULT**: The calculated sum value

## Troubleshooting

### Issue: "Column not found in table" Error

**Cause**: Column name doesn't match or doesn't exist

**Solution**:
1. Verify column names are spelled correctly (case-sensitive)
2. Check column names match table headers exactly
3. Use "Get Table Rows" to inspect table structure

### Issue: Unexpected Sum Result

**Cause**: Non-numeric values in columns

**Solution**:
1. Check for non-numeric values in columns
2. Remember non-numeric values are treated as 0
3. Verify data types in Excel

### Issue: Sum is Zero

**Cause**: Table has no data rows or all values are non-numeric

**Solution**:
1. Verify table contains data rows
2. Check columns contain numeric values
3. Use "Get Table Rows" to inspect data

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Range Sum](pages/RangeSum.md) - Sum range values
- [Get Table Rows](pages/GetTableRows.md) - Get all table data

