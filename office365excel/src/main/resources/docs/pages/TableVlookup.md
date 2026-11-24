# Table Vlookup

## Overview

Performs a VLOOKUP operation on an Excel table in a personal OneDrive file, searching for a value in the first column and returning a corresponding value from a specified column. This is the programmatic equivalent of Excel's VLOOKUP function.

## Request Details

- **Area**: Table
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | The ID of the Excel file containing the table | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Lookup Value | Text | Yes | The value to search for in the first column of the table | "EMP001" |
| Table Name | Text | Yes | The name of the table to perform VLOOKUP on | "EmployeeTable" |
| Column Index | Number | Yes | The column number to return the value from (1-based index) | 3 |
| Exact Match | Boolean | Yes | True for exact match, False for approximate match | true |

### Parameter Details

#### File Id
- **Format**: Unique identifier for the Excel file in OneDrive
- **How to obtain**: Use "Get All Excel Files For Name" catalog request
- **Validation**: Must be a valid Excel file ID accessible by the authenticated user
- **Example**: `01C8MOQGI3YRBDXAQERECWE3W4XCF`

#### Lookup Value
- **Format**: Text or numeric value to search for
- **Data Types**: Supports text and numbers only
- **Case Sensitivity**: Case-insensitive for text values
- **Matching**: Searches in the first column of the table
- **Example**: `"EMP001"`, `"John Doe"`, `"12345"`

> **📝 Note**: Lookup value supports text and numbers only. Dates and boolean values should be converted to text.

#### Table Name
- **Format**: Name of the table as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Table must exist in the Excel file
- **Example**: `Table1`, `EmployeeData`, `SalesRecords`

#### Column Index
- **Format**: Positive integer representing column position
- **Range**: 1 to number of columns in the table
- **1-Based**: First column is 1 (not 0)
- **Validation**: Must not exceed the number of columns in the table
- **Example**: `2` (second column), `5` (fifth column)

#### Exact Match
- **Format**: Boolean value (true/false)
- **True**: Finds exact match only (recommended for most cases)
- **False**: Finds approximate match (requires sorted data)
- **Default Behavior**: Use `true` for exact matching
- **Example**: `true`, `false`

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Lookup Result | FreeForm | Object containing the lookup key and corresponding value |

### Lookup Result Structure

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| Lookup Key | Text | The value that was searched for | "EMP001" |
| Lookup Value | Text | The corresponding value from the specified column | "John Doe" |

### Output Example

```json
{
  "Lookup Result": {
    "Lookup Key": "EMP001",
    "Lookup Value": "John Doe"
  }
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Lookup Value is empty or null | "Lookup Value is required" | Provide a value to search for |
| Table Name is empty or null | "Table Name is required" | Provide a valid table name |
| Column Index is null or ≤ 0 | "Column Index must be a positive number" | Provide a valid column index (≥ 1) |
| Exact Match is null | "Exact Match is required" | Provide true or false |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Table not found | "Table '{tableName}' not found in the file" | Verify table name (case-sensitive) |
| Column Index exceeds table columns | "Column Index {index} exceeds table column count" | Use valid column index within table range |
| Lookup value not found (exact match) | "Lookup value not found in table" | Verify lookup value exists in first column |
| Table is empty | "Table contains no data" | Ensure table has at least one data row |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters
- Column Index is zero or negative
- Invalid data types

**Error Message**: `"File ID is required"`, `"Column Index must be a positive number"`

**Resolution**:
1. Verify all required parameters are provided
2. Ensure Column Index is a positive integer
3. Check Exact Match is boolean (true/false)

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Table doesn't exist in the file
- Column Index exceeds number of columns
- Lookup value not found (when Exact Match = true)
- Table has no data rows
- File not accessible

**Error Message**: `"Table '{tableName}' not found"`, `"Lookup value not found in table"`

**Resolution**:
1. Verify table name is correct (case-sensitive)
2. Check Column Index is within valid range
3. Ensure lookup value exists in first column
4. Verify table contains data rows
5. Check user has read permissions

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to perform VLOOKUP"`, `"Microsoft Graph API error: {details}"`

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
2. Verify user has Files.ReadWrite permission
3. Check user has read access to the file

## Usage Examples

### Example 1: Employee Lookup by ID

**Scenario**: Find employee name by employee ID

**Table Structure**:
| Employee ID | Name | Department | Salary |
|-------------|------|------------|--------|
| EMP001 | John Doe | Engineering | 75000 |
| EMP002 | Jane Smith | Marketing | 65000 |
| EMP003 | Bob Johnson | Sales | 70000 |

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Lookup Value: "EMP001"
Table Name: "EmployeeTable"
Column Index: 2
Exact Match: true
```

**Output**:
```json
{
  "Lookup Result": {
    "Lookup Key": "EMP001",
    "Lookup Value": "John Doe"
  }
}
```

**Result**: Successfully found employee name "John Doe" for ID "EMP001"

### Example 2: Product Price Lookup

**Scenario**: Find product price by product code

**Table Structure**:
| Product Code | Product Name | Price | Stock |
|--------------|--------------|-------|-------|
| PRD001 | Widget A | 29.99 | 150 |
| PRD002 | Widget B | 39.99 | 200 |
| PRD003 | Widget C | 49.99 | 100 |

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Lookup Value: "PRD002"
Table Name: "ProductTable"
Column Index: 3
Exact Match: true
```

**Output**:
```json
{
  "Lookup Result": {
    "Lookup Key": "PRD002",
    "Lookup Value": "39.99"
  }
}
```

**Result**: Successfully found price "39.99" for product "PRD002"

### Example 3: Lookup Value Not Found

**Scenario**: Search for non-existent employee ID

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Lookup Value: "EMP999"
Table Name: "EmployeeTable"
Column Index: 2
Exact Match: true
```

**Error**: LOGIC_ERROR - "Lookup value not found in table"

**Result**: No matching record found for "EMP999"

### Example 4: Multi-Column Lookup

**Scenario**: Get department for an employee

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Lookup Value: "EMP001"
Table Name: "EmployeeTable"
Column Index: 3
Exact Match: true
```

**Output**:
```json
{
  "Lookup Result": {
    "Lookup Key": "EMP001",
    "Lookup Value": "Engineering"
  }
}
```

**Result**: Successfully found department "Engineering" for employee "EMP001"

## Business Rules

1. **First Column Search**: Always searches in the first column of the table
2. **1-Based Indexing**: Column Index starts at 1 (first column = 1)
3. **Case-Insensitive**: Text matching is case-insensitive
4. **Single Result**: Returns only the first matching value
5. **Exact Match Recommended**: Use Exact Match = true for most scenarios
6. **Approximate Match**: Requires data to be sorted in ascending order (rarely used)
7. **Data Types**: Lookup value supports text and numbers only
8. **Empty Cells**: Empty cells in lookup column are skipped

## Limitations

1. **First Column Only**: Can only search in the first column of the table
2. **Single Value**: Returns only one value, not multiple matches
3. **No Wildcards**: Does not support wildcard matching (*, ?)
4. **Text and Numbers Only**: Lookup value must be text or number (no dates, booleans)
5. **No Formula**: Returns value only, not formulas
6. **Personal OneDrive Only**: For shared drives, use "VLOOKUP in Table" from Shared Area
7. **Table Required**: Cannot perform VLOOKUP on regular ranges (must be an Excel table)

## Best Practices

### 1. Use Exact Match
Always use `Exact Match = true` unless you specifically need approximate matching with sorted data.

### 2. Validate Lookup Value Exists
Check if the lookup value exists before performing VLOOKUP to handle "not found" scenarios gracefully.

### 3. Use Meaningful Column Indices
Document which column index corresponds to which data field for maintainability.

### 4. Handle Not Found Errors
Implement error handling for cases where the lookup value doesn't exist in the table.

### 5. Verify Table Structure
Ensure the table structure matches your expectations before performing lookups.

### 6. Cache Table Metadata
If performing multiple lookups on the same table, cache table structure information.

## Common Use Cases

### 1. Employee Information Retrieval
```
Scenario: Look up employee details by employee ID
Action: VLOOKUP with employee ID to get name, department, or salary
Result: Employee information retrieved for display or processing
```

### 2. Product Catalog Lookup
```
Scenario: Find product details by product code
Action: VLOOKUP with product code to get price, description, or stock level
Result: Product information available for order processing
```

### 3. Customer Data Retrieval
```
Scenario: Get customer information by customer ID
Action: VLOOKUP with customer ID to retrieve contact details
Result: Customer data available for communication or reporting
```

### 4. Reference Data Lookup
```
Scenario: Translate codes to descriptions (e.g., country code to country name)
Action: VLOOKUP with code to get corresponding description
Result: Human-readable description obtained for display
```

## Related Catalog Requests

- [Vlookup](pages/Vlookup.md) - Sheet Area version for VLOOKUP on ranges
- [VLOOKUP in Table](pages/VLOOKUPInTable.md) - Shared Area version for shared drive files
- [Get Rows](pages/GetRows.md) - Retrieve all rows from a table
- [Get Column Values for Keys](pages/GetColumnValuesForKeys.md) - Get multiple column values for keys
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Find File ID

## Technical Implementation

### Helper Class
- **Class**: TableOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation and table structure verification
- **Service**: Delegates to WorkbookService for table operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/items/{fileId}/workbook/tables/{tableName}/rows`
- **Processing**: Client-side VLOOKUP logic on retrieved table data
- **Permissions Required**: Files.ReadWrite

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **VLOOKUP_MATCH_FOUND**: Logged when match is found
- **VLOOKUP_NO_MATCH**: Logged when no match is found
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Lookup value not found" Error

**Cause**: Value doesn't exist in first column of table

**Solution**:
1. Verify lookup value is spelled correctly
2. Check for extra spaces or special characters
3. Ensure value exists in the first column
4. Verify case matches (though matching is case-insensitive)

### Issue: "Column Index exceeds table column count"

**Cause**: Column Index is larger than number of columns in table

**Solution**:
1. Count the number of columns in your table
2. Ensure Column Index is within valid range (1 to column count)
3. Use "Get Rows" to inspect table structure

### Issue: Wrong Value Returned

**Cause**: Incorrect Column Index specified

**Solution**:
1. Verify Column Index points to the correct column
2. Remember Column Index is 1-based (first column = 1)
3. Count columns carefully, including the lookup column

### Issue: Approximate Match Not Working

**Cause**: Data not sorted or Exact Match set incorrectly

**Solution**:
1. For approximate match, ensure first column is sorted in ascending order
2. Set Exact Match = false for approximate matching
3. Consider using Exact Match = true instead (recommended)

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get Rows](pages/GetRows.md) - Retrieve all table rows
- [Vlookup](pages/Vlookup.md) - VLOOKUP on sheet ranges

