# Search

## Overview

Searches for a specific value within a worksheet of an Excel file stored in a shared drive and returns the cell IDs (addresses) where the value is found.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the sheet | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | The name of the sheet to search in | "Sheet1" |
| Data to Search | Text | Yes | The text or value to search for in the sheet | "John Doe" |

### Parameter Details

#### Drive Id
- **Format**: Unique identifier for the shared drive
- **How to obtain**: Use "Get All Shared Drives Information" catalog request
- **Example**: `b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8`

#### File Id
- **Format**: Unique identifier for the Excel file
- **How to obtain**: Use "Get Shared Excel Files" catalog request
- **Example**: `01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR`

#### Sheet Name
- **Format**: Name of the worksheet as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Example**: `Sheet1`, `Sales Data`

#### Data to Search
- **Format**: Text or numeric value to search for
- **Case Sensitivity**: Case-insensitive search
- **Exact Match**: Searches for exact cell value matches
- **Example**: `"John Doe"`, `"12345"`, `"Product A"`

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell Ids | List&lt;Text&gt; | List of cell addresses where the search value was found |

### Output Example

```json
{
  "Cell Ids": ["A2", "C5", "E10"]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty or null | "Drive ID is required" | Provide a valid Drive ID |
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| Data to Search is empty or null | "Data to Search is required" | Provide a search value |
| Drive not found | "Drive not found or inaccessible" | Verify Drive ID and permissions |
| File not found | "File not found or inaccessible" | Verify File ID and permissions |
| Sheet not found | "Sheet '{sheetName}' not found" | Verify sheet name (case-sensitive) |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing required parameters
**Resolution**: Provide all required parameters

### Logic Errors (LOGIC_ERROR)
**Cause**: Sheet doesn't exist or file not accessible
**Resolution**: Verify sheet name and file accessibility

### System Errors (SYSTEM_ERROR)
**Cause**: Microsoft Graph API failures
**Resolution**: Retry operation, check connectivity

### Authorization Errors
**Cause**: Authentication or permission issues
**Resolution**: Re-authenticate, verify permissions

## Usage Examples

### Example 1: Find Employee Name

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Employees"
Data to Search: "John Doe"
```

**Output**:
```json
{
  "Cell Ids": ["B2"]
}
```

**Result**: "John Doe" found in cell B2

### Example 2: Find Multiple Occurrences

**Input**:
```
Data to Search: "Active"
```

**Output**:
```json
{
  "Cell Ids": ["D2", "D5", "D8", "D12"]
}
```

**Result**: "Active" found in 4 cells

### Example 3: Value Not Found

**Input**:
```
Data to Search: "NonExistent"
```

**Output**:
```json
{
  "Cell Ids": []
}
```

**Result**: No matches found

## Business Rules

1. **Case-Insensitive**: Search is not case-sensitive
2. **Exact Match**: Only exact cell value matches are returned
3. **All Occurrences**: Returns all cells containing the value
4. **Order**: Results are returned in row-then-column order
5. **Empty Cells**: Empty cells are not searched

## Limitations

1. **Exact Match Only**: Does not support partial or wildcard matching
2. **Single Sheet**: Searches only one sheet per request
3. **No Formulas**: Searches calculated values, not formula text
4. **Shared Drives Only**: For shared drive files only
5. **Performance**: Large sheets may take longer to search

## Best Practices

### 1. Use Specific Search Terms
Use specific values to reduce search time and result set size.

### 2. Handle Empty Results
Check if Cell Ids list is empty before processing results.

### 3. Validate Cell Addresses
Verify returned cell addresses before using them in other operations.

## Common Use Cases

### 1. Find Data Location
```
Scenario: Locate where specific data appears in sheet
Action: Search for value
Result: Cell addresses where value is found
```

### 2. Data Validation
```
Scenario: Check if value exists in sheet
Action: Search for value and check if results are empty
Result: Confirmation of value presence
```

### 3. Update Multiple Cells
```
Scenario: Find all cells with specific value to update
Action: Search for value, then update each cell
Result: All matching cells updated
```

## Related Catalog Requests

- [Get Cell Value](pages/GetCellValue.md) - Read cell values
- [Add Data to Provided Cell](pages/AddDataToProvidedCell.md) - Write to cells
- [Get Rows for Range](pages/GetRowsForRange.md) - Get range data
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get File ID

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Service**: Delegates to WorksheetService for search operations

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/usedRange`
- **Processing**: Client-side search on retrieved data
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **MATCHES_FOUND_COUNT**: Number of matching cells found

## Troubleshooting

### Issue: No Results Found

**Cause**: Value doesn't exist or search term doesn't match exactly

**Solution**:
1. Verify search value matches cell content exactly
2. Check for extra spaces or special characters
3. Remember search is case-insensitive but must be exact match

### Issue: Slow Search Performance

**Cause**: Large sheet with many cells

**Solution**:
1. This is expected for large sheets
2. Consider limiting search to specific ranges if possible

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get Cell Value](pages/GetCellValue.md) - Read cell values
- [Add Data to Provided Cell](pages/AddDataToProvidedCell.md) - Write to cells

