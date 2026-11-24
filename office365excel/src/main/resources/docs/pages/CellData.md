# Cell Data

## Overview

Retrieves comprehensive information about a specific cell in a worksheet of an Excel file stored in a shared drive, including the cell ID and its current value.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | The ID of the Excel file containing the sheet | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | The name of the sheet containing the cell | "Sheet1" |
| Cell Id | Text | Yes | The ID of the cell to retrieve data from | "A1" |

### Parameter Details

#### Cell Id
- **Format**: Excel cell reference (column letter + row number)
- **Pattern**: `[A-Z]+[0-9]+` (e.g., A1, AA100, Z999)
- **Case Sensitivity**: Not case-sensitive (A1 = a1)
- **Example**: `A1`, `B5`, `AA10`

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell Information | FreeForm | Object containing cell ID and value |

### Cell Information Structure

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| Cell ID | Text | The identifier of the cell | "A1" |
| Value | Text | The current value in the cell | "Hello World" |

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
| Cell Id is empty or null | "Cell ID is required" | Provide a valid cell ID |
| Invalid cell ID format | "Invalid cell ID format" | Use valid Excel cell reference |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing or invalid parameters
**Resolution**: Provide all required parameters with valid formats

### Logic Errors (LOGIC_ERROR)
**Cause**: Sheet doesn't exist or file not accessible
**Resolution**: Verify sheet name and file accessibility

### System Errors (SYSTEM_ERROR)
**Cause**: Microsoft Graph API failures
**Resolution**: Retry operation

### Authorization Errors
**Cause**: Authentication or permission issues
**Resolution**: Re-authenticate, verify permissions

## Usage Examples

### Example 1: Get Cell Value

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Sheet1"
Cell Id: "A2"
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

### Example 2: Get Numeric Value

**Input**:
```
Cell Id: "B5"
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

### Example 3: Get Empty Cell

**Input**:
```
Cell Id: "Z99"
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

## Business Rules

1. **Formula Values**: Returns calculated result, not formula text
2. **Data Types**: All values returned as text strings
3. **Empty Cells**: Empty cells return empty string value
4. **No Formatting**: Cell formatting not included

## Limitations

1. **Single Cell Only**: Cannot retrieve multiple cells
2. **No Formatting**: Cell formatting not returned
3. **No Formula Text**: Cannot retrieve formula itself
4. **Shared Drives Only**: For shared drive files only

## Best Practices

### 1. Validate Cell References
Validate cell ID format before making request.

### 2. Handle Empty Cells
Check for empty string values appropriately.

### 3. Type Conversion
Convert string values to appropriate types as needed.

## Common Use Cases

### 1. Read Configuration Value
```
Scenario: Read configuration from specific cell
Action: Get cell data
Result: Configuration value retrieved
```

### 2. Validate Data Entry
```
Scenario: Check if cell contains expected value
Action: Retrieve cell value and compare
Result: Validation passed or failed
```

## Related Catalog Requests

- [Get Cell Value](pages/GetCellValue.md) - Sheet Area version
- [Add Data to Provided Cell](pages/AddDataToProvidedCell.md) - Write to cell
- [Search](pages/Search.md) - Find cells with specific values
- [Get Rows for Range](pages/GetRowsForRange.md) - Get multiple cells

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Service**: Delegates to WorksheetService

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{cellId}')`
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error

## Troubleshooting

### Issue: "Invalid cell ID format" Error

**Solution**: Use format: Column letter(s) + Row number (e.g., A1, AA10)

### Issue: Empty Value Returned

**Solution**: Verify cell actually contains data in Excel

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)
- [Get Cell Value](pages/GetCellValue.md)
- [Add Data to Provided Cell](pages/AddDataToProvidedCell.md)

