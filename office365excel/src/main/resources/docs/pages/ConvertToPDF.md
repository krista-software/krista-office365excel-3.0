# Convert To PDF

## Overview

Converts an Excel file stored in a shared drive to PDF format while preserving the original formatting.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | The ID of the Excel file to convert to PDF | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Drive Id | Text | Yes | The ID of the drive where the Excel file is located | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| PDF File | File | The converted PDF file |

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| File Id is empty or null | "File Id is required" | Provide a valid File Id |
| Drive Id is empty or null | "Drive Id is required" | Provide a valid Drive Id |
| File not found | "File not found or inaccessible" | Verify File Id and Drive Id |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing required parameters
**Resolution**: Provide all required parameters

### Logic Errors (LOGIC_ERROR)
**Cause**: File doesn't exist or not accessible
**Resolution**: Verify file exists and is accessible

### System Errors (SYSTEM_ERROR)
**Cause**: Conversion failure or API issues
**Resolution**: Retry operation

## Usage Examples

### Example 1: Convert Excel to PDF

**Input**:
```
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
```

**Output**:
```
PDF File: [Binary PDF file data]
```

**Result**: Excel file converted to PDF with same formatting

## Business Rules

1. **Preserves Formatting**: Original Excel formatting maintained
2. **All Sheets**: All worksheets included in PDF
3. **Read-Only**: Original Excel file not modified
4. **Same Name**: PDF has same base name as Excel file

## Limitations

1. **File Size**: Large files may take longer to convert
2. **Complex Formatting**: Some complex formatting may not convert perfectly
3. **Shared Drives Only**: For shared drive files only
4. **No Customization**: Cannot customize PDF output settings

## Best Practices

### 1. Verify File Format
Ensure file is a valid Excel file before conversion.

### 2. Handle Large Files
Implement timeout handling for large file conversions.

### 3. Save PDF
Save the returned PDF file to desired location.

## Common Use Cases

### 1. Report Distribution
```
Scenario: Convert Excel report to PDF for distribution
Action: Convert file to PDF
Result: PDF version of report created
```

### 2. Archive Documents
```
Scenario: Create PDF archive of Excel files
Action: Convert to PDF for long-term storage
Result: PDF archive created
```

## Related Catalog Requests

- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Find files to convert
- [Get Shared Excel File](pages/GetSharedExcelFile.md) - Get specific file

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/items/{fileId}/content?format=pdf`
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful conversion
- **CONVERSION_DURATION**: Time taken to convert

## Troubleshooting

### Issue: Conversion Takes Long Time

**Cause**: Large or complex Excel file

**Solution**: This is expected for large files, wait for completion

### Issue: Formatting Issues in PDF

**Cause**: Complex Excel formatting

**Solution**: Simplify Excel formatting if possible

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md)

