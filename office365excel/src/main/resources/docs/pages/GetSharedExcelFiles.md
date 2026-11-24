# Get Shared Excel Files

## Overview

Retrieves a list of Excel files (including CSV, TSV, XLS, and XLSX formats) from shared drives accessible to the authenticated user. This request provides essential file metadata needed for subsequent operations on shared Excel files.

> **Root-level files only**: When a Drive ID is provided, this request returns only the files located directly in the **root folder** of that drive. Files stored inside sub-folders are **not** included. To search recursively across all folders and sub-folders within a drive, use [Search Excel Files in Drive](pages/SearchExcelFilesInDrive.md) instead.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

This catalog request does not require any input parameters.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Files | List&lt;Entity(Excel File)&gt; | List of Excel files from all accessible shared drives |

### Excel File Entity Structure

Each Excel File entity contains the following properties:

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| Drive ID | Text | Unique identifier for the drive containing the file | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File ID | Text | Unique identifier for the Excel file | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| File Name | Text | Name of the Excel file | "Book.xlsx" |
| Web URL | Text | Web URL to access the file in browser | "https://contoso.sharepoint.com/sites/Marketing/Book.xlsx" |
| Created Date | Text | Date when the file was created | "2024-01-15T10:30:00Z" |
| Modified Date | Text | Date when the file was last modified | "2024-01-20T14:45:00Z" |

### Output Example

```json
{
  "Files": [
    {
      "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR",
      "File Name": "Sales_Report.xlsx",
      "Web URL": "https://contoso.sharepoint.com/sites/Sales/Sales_Report.xlsx",
      "Created Date": "2024-01-15T10:30:00Z",
      "Modified Date": "2024-01-20T14:45:00Z"
    },
    {
      "Drive ID": "b!xYzABm0EY0WK0z3E4xcAH5Ufuu38G9",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROK",
      "File Name": "Employee_Data.csv",
      "Web URL": "https://contoso.sharepoint.com/sites/HR/Employee_Data.csv",
      "Created Date": "2024-01-10T09:00:00Z",
      "Modified Date": "2024-01-22T16:20:00Z"
    }
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| User not authenticated | "Authentication required" | Ensure user is properly authenticated |
| Invalid access token | "Invalid or expired access token" | Re-authenticate to obtain new token |
| Insufficient permissions | "Insufficient permissions to access shared files" | Verify Files.ReadWrite.All permission is granted |

## Error Handling

### Input Errors (INPUT_ERROR)

This request has no input parameters, so input errors are not applicable.

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic or permission issues

**Common Scenarios**:
- User has no shared drives accessible
- User has no Excel files in accessible shared drives
- User account lacks Office 365 license
- Conditional access policies blocking access

**Error Message**: `"No shared Excel files found"` or `"Access denied to shared drives"`

**Resolution**:
1. Verify user has access to at least one shared drive
2. Check that shared drives contain Excel files
3. Ensure user has valid Office 365 license
4. Review conditional access policies

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to retrieve shared Excel files"` or `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions (Files.ReadWrite.All not granted)
- User doesn't have access to any shared drives

**Error Message**: `"Unauthorized access"` or `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify Files.ReadWrite.All permission is granted
3. Ensure admin consent was provided for the permission
4. Check user has access to at least one shared drive

## Usage Examples

### Example 1: List All Shared Excel Files

**Scenario**: Retrieve all Excel files from shared drives for user selection

**Input**: None

**Output**:
```json
{
  "Files": [
    {
      "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR",
      "File Name": "Q1_Budget.xlsx",
      "Web URL": "https://contoso.sharepoint.com/sites/Finance/Q1_Budget.xlsx"
    },
    {
      "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROS",
      "File Name": "Inventory.csv",
      "Web URL": "https://contoso.sharepoint.com/sites/Operations/Inventory.csv"
    }
  ]
}
```

**Result**: Two Excel files retrieved from shared drives

### Example 2: Find File IDs for Subsequent Operations

**Scenario**: Get File ID and Drive ID to use in "Add Rows to Table" request

**Input**: None

**Output**:
```json
{
  "Files": [
    {
      "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR",
      "File Name": "Sales_Data.xlsx",
      "Web URL": "https://contoso.sharepoint.com/sites/Sales/Sales_Data.xlsx"
    }
  ]
}
```

**Next Step**: Use Drive ID and File ID in subsequent catalog requests

**Result**: File metadata obtained for use in other operations

### Example 3: Multiple File Formats

**Scenario**: Retrieve Excel files in various formats (XLSX, CSV, TSV, XLS)

**Input**: None

**Output**:
```json
{
  "Files": [
    {
      "File Name": "Report.xlsx",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
    },
    {
      "File Name": "Data.csv",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROS"
    },
    {
      "File Name": "Export.tsv",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROT"
    },
    {
      "File Name": "Legacy.xls",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROU"
    }
  ]
}
```

**Result**: All supported Excel file formats retrieved

### Example 4: No Excel Files Available

**Scenario**: User has shared drives but no Excel files

**Input**: None

**Output**:
```json
{
  "Files": []
}
```

**Result**: Empty list indicates no Excel files in accessible shared drives

## Business Rules

1. **Supported Formats**: Returns files with extensions: .xlsx, .xls, .csv, .tsv
2. **Permission-Based**: Only files the authenticated user has access to are returned
3. **All Shared Drives**: Searches across all accessible shared drives
4. **Real-Time Data**: Results reflect current file availability and permissions
5. **No Filtering**: All Excel files from all shared drives are returned
6. **Metadata Only**: Only file metadata is returned, not file contents
7. **Root-Level Only**: When a Drive ID is provided, only files in the root folder of that drive are returned — files inside sub-folders are not included

## Limitations

1. **Personal OneDrive Excluded**: User's personal OneDrive files are not included
2. **Shared Drives Only**: Only returns files from shared drives (OneDrive shared folders, SharePoint)
3. **No Content**: File contents are not retrieved, only metadata
4. **No Filtering**: Cannot filter by drive, folder, or file name in this request
5. **Large Result Sets**: In organizations with many files, response may be large
6. **Permission Changes**: Recent permission changes may take a few minutes to reflect
7. **Deleted Files**: Recently deleted files may still appear temporarily
8. **Root Folder Only (Drive ID)**: When a Drive ID is provided, only files at the root level of the drive are returned. Files stored in any sub-folder within that drive will **not** appear in the results. Use [Search Excel Files in Drive](pages/SearchExcelFilesInDrive.md) for recursive sub-folder search

## Best Practices

### 1. Cache File List
Store file list locally to avoid repeated API calls. Refresh periodically or when needed.

### 2. Handle Empty Results
Always check if the Files list is empty before attempting to access file properties.

### 3. Use File ID for Operations
Always use the File ID (not the name) for subsequent operations, as names can change.

### 4. Display User-Friendly Names
Show File Name to users for selection, but use File ID and Drive ID internally.

### 5. Filter Client-Side
If you need to filter files, retrieve all files and filter on the client side.

### 6. Monitor Performance
For organizations with many files, consider implementing pagination or filtering strategies.

## Common Use Cases

### 1. File Selection Interface
```
Scenario: Display a dropdown list of available Excel files for user selection
Action: Call Get Shared Excel Files and populate dropdown with File Names
Result: User can select a file to work with
```

### 2. Automated File Processing
```
Scenario: Process all Excel files across shared drives
Action: Retrieve all files, then iterate through each to perform operations
Result: Comprehensive file processing across organization
```

### 3. File Discovery
```
Scenario: Help users discover what Excel files they have access to
Action: Display all accessible files with metadata
Result: Users can see all available Excel resources
```

### 4. File Validation
```
Scenario: Verify a specific file exists before attempting operations
Action: Retrieve all files and check if target file is in the list
Result: Confirm file exists and is accessible
```

## Related Catalog Requests

- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Get drive information first
- [Search Excel Files in Drive](pages/SearchExcelFilesInDrive.md) - Recursive alternative: searches all folders and sub-folders within a drive by file format
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Add data to tables in shared files
- [Get Table Rows](pages/GetTableRows.md) - Retrieve data from tables in shared files
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add data to sheets in shared files
- [Get Excel File Information](pages/GetExcelFileInformation.md) - Personal OneDrive version

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog
- **Validation**: Minimal validation (authentication check only)
- **Service**: Delegates to DriveService.getSharedExcelFilesInformation()

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drives` and `GET /drives/{driveId}/root/children`
- **Filter**: Filters for Excel file extensions (.xlsx, .xls, .csv, .tsv)
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration
- **FILES_RETRIEVED_COUNT**: Number of files returned

## Troubleshooting

### Issue: Empty Files List

**Cause**: No Excel files in accessible shared drives

**Solution**:
1. Verify user has access to shared drives
2. Check that shared drives contain Excel files (.xlsx, .xls, .csv, .tsv)
3. Ensure files are not in personal OneDrive (use different request)
4. Verify user permissions to access files

### Issue: "Insufficient Permissions" Error

**Cause**: Files.ReadWrite.All permission not granted

**Solution**:
1. Navigate to Microsoft Entra ID application
2. Add Files.ReadWrite.All delegated permission
3. Grant admin consent
4. Re-authenticate user

### Issue: Missing Files

**Cause**: User doesn't have access to all files or recent changes

**Solution**:
1. Verify user has read permissions to the files
2. Check file sharing settings
3. Wait a few minutes for permission changes to propagate
4. Ensure files are in shared drives, not personal OneDrive

### Issue: Files in Sub-Folders Not Returned

**Cause**: This request only lists files in the root folder of the specified drive — sub-folder contents are not traversed

**Solution**:
1. Use [Search Excel Files in Drive](pages/SearchExcelFilesInDrive.md) instead — it performs a recursive search across all folders and sub-folders within a drive
2. Confirm the file you are looking for is not stored inside a sub-folder

### Issue: Large Response Time

**Cause**: Many files across multiple drives

**Solution**:
1. This is expected behavior for large organizations
2. Consider caching results
3. Implement client-side filtering
4. Use pagination strategies if needed

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Get drive information
- [Get Excel File Information](pages/GetExcelFileInformation.md) - Personal OneDrive version

