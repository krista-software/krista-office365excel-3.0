# Get Shared Excel File

## Overview

Retrieves a specific Excel file from a shared drive folder by file name and optional folder name.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Shared Folder Name | Text | No | The name of the folder containing the Excel file | "Reports" |
| Shared File Name | Text | Yes | The name of the Excel file to retrieve | "Sales_Report.xlsx" |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Shared Excel File | Entity(Excel File) | The Excel file entity with file details |

### Excel File Entity Structure

| Property | Type | Description |
|----------|------|-------------|
| File ID | Text | Unique identifier for the file |
| File Name | Text | Name of the Excel file |
| Drive ID | Text | ID of the drive containing the file |
| Web URL | Text | URL to access the file in browser |

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Shared File Name is empty or null | "Shared File Name is required" | Provide a valid file name |
| File not found | "File not found" | Verify file name and folder |

## Error Handling

### Input Errors (INPUT_ERROR)
**Cause**: Missing file name
**Resolution**: Provide file name

### Logic Errors (LOGIC_ERROR)
**Cause**: File doesn't exist in specified location
**Resolution**: Verify file name and folder name

## Usage Examples

### Example 1: Get File from Root

**Input**:
```
Shared Folder Name: null
Shared File Name: "Sales_Report.xlsx"
```

**Output**:
```json
{
  "Shared Excel File": {
    "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR",
    "File Name": "Sales_Report.xlsx",
    "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
    "Web URL": "https://..."
  }
}
```

### Example 2: Get File from Specific Folder

**Input**:
```
Shared Folder Name: "Reports"
Shared File Name: "Q1_Report.xlsx"
```

**Output**:
```json
{
  "Shared Excel File": {
    "File ID": "01ABCDEF123456789",
    "File Name": "Q1_Report.xlsx",
    "Drive ID": "b!xyz123",
    "Web URL": "https://..."
  }
}
```

## Business Rules

1. **Case-Sensitive**: File and folder names are case-sensitive
2. **Exact Match**: Requires exact file name match
3. **First Match**: Returns first matching file if duplicates exist

## Limitations

1. **Single File**: Returns only one file
2. **Exact Name**: No wildcard or partial matching
3. **Shared Drives Only**: For shared drive files only

## Best Practices

### 1. Use Exact Names
Provide exact file and folder names.

### 2. Handle Not Found
Implement error handling for missing files.

## Common Use Cases

### 1. Locate Specific File
```
Scenario: Find file by name for processing
Action: Get file with exact name
Result: File details retrieved
```

## Related Catalog Requests

- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get all files
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Get drives

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/root/search(q='{fileName}')`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Get Shared Excel Files](pages/GetSharedExcelFiles.md)
- [Extension Configuration](pages/ExtensionConfiguration.md)

