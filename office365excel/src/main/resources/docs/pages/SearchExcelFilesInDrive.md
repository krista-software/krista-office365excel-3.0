# Search Excel Files in Drive

## Overview

Searches for all files of a selected format recursively across all folders in a specified shared drive using the Microsoft Graph search API. Unlike **Get Shared Excel Files** which scans top-level shared items, this request performs a deep recursive traversal of every folder inside a known drive and returns all matching files — including files stored in nested sub-folders.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ❌ No - Does not use the allow-retry mechanism; failed extension searches are logged and skipped

## Input Parameters

| Parameter Name | Type | Required | Description |
|----------------|------|----------|-------------|
| Drive Id | Text | ✅ Yes | Unique identifier of the drive to search in. Obtain this from **Get All Shared Drives Information** |
| File Format | PickOne | ✅ Yes | File extension to search for. One of: `.xlsx`, `.xls`, `.csv`, `.tsv` |

### Input Parameter Details

#### Drive Id
The Drive ID uniquely identifies a SharePoint document library or OneDrive shared folder. It typically begins with `b!` for SharePoint drives.

**Example**: `b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3`

> **Tip**: Use **Get All Shared Drives Information** to discover available Drive IDs before calling this request.

#### File Format (PickOne)
Selects the file extension to search for within the drive. Only one format can be selected per request.

| Value | Description |
|-------|-------------|
| `.xlsx` | Excel Workbook (Office 2007+) |
| `.xls` | Excel 97–2003 Workbook (legacy) |
| `.csv` | Comma-Separated Values |
| `.tsv` | Tab-Separated Values |

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Files | List&lt;Entity(Excel File)&gt; | All files of the selected format found recursively in the drive |

### Excel File Entity Structure

Each Excel File entity contains the following properties:

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| Drive ID | Text | Unique identifier of the drive containing the file | `b!jxUeN0DgA02sUFcc...` |
| File ID | Text | Unique identifier of the file | `01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR` |
| File Name | Text | Name of the file including extension | `Sales_Report.xlsx` |
| Web URL | Text | Browser URL to open the file directly | `https://contoso.sharepoint.com/sites/Sales/Sales_Report.xlsx` |

### Output Example

```json
{
  "Files": [
    {
      "Drive ID": "b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR",
      "File Name": "Sales_Report.xlsx",
      "Web URL": "https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Sales_Report.xlsx"
    },
    {
      "Drive ID": "b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROS",
      "File Name": "Q1_Budget.xlsx",
      "Web URL": "https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Archive/Q1_Budget.xlsx"
    }
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Drive Id is empty | `"Drive Id is required"` | Provide a valid Drive ID |
| File Format not selected | `"File Format is required"` | Select one of the supported file formats |
| Drive not accessible | `itemNotFound (404)` | Verify the Drive ID is correct and the user has access |
| Authentication missing | `"Authentication required"` | Ensure the user is properly authenticated |

## Error Handling

### Input Errors (INPUT_ERROR)

**Drive Id blank or missing**

**Resolution**: Provide the Drive ID obtained from **Get All Shared Drives Information**.

### Logic Errors (LOGIC_ERROR)

**Cause**: Drive is not accessible or contains no files of the selected format.

**Common Scenarios**:
- Drive ID points to a drive the user cannot access
- No files with the selected extension exist in the drive
- User account lacks Office 365 license

**Error Message**: Returns an empty `Files` list — no exception is thrown

**Resolution**:
1. Verify the Drive ID is correct
2. Confirm the drive contains files with the selected extension
3. Ensure the user has at least read access to the drive

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API failures during search.

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Drive not found (404 — `itemNotFound`)

**Error Message**: Search failure is logged per extension; method returns empty list rather than throwing

**Resolution**:
1. Verify the Drive ID is correct and accessible
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Use `/drives/{driveId}` endpoint — **not** `/me/drives/{driveId}` — for SharePoint drives (this is already handled by the implementation)

### Authorization Errors

**Cause**: Authentication or permission issues.

**Common Scenarios**:
- Access token expired
- `Files.ReadWrite.All` permission not granted
- User does not have access to the specified drive

**Resolution**:
1. Re-authenticate if token expired
2. Verify `Files.ReadWrite.All` delegated permission is granted
3. Ensure admin consent was provided for the permission

## Usage Examples

### Example 1: Search for All XLSX Files in a SharePoint Drive

**Scenario**: Find all Excel workbooks in a SharePoint document library including sub-folders

**Input**:
```
Drive Id:    b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3
File Format: .xlsx
```

**Output**:
```json
{
  "Files": [
    {
      "Drive ID": "b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR",
      "File Name": "Annual_Report.xlsx",
      "Web URL": "https://contoso.sharepoint.com/sites/Finance/Shared%20Documents/Annual_Report.xlsx"
    },
    {
      "Drive ID": "b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROS",
      "File Name": "Budget_2025.xlsx",
      "Web URL": "https://contoso.sharepoint.com/sites/Finance/Shared%20Documents/Reports/Budget_2025.xlsx"
    }
  ]
}
```

**Result**: Two `.xlsx` files found — one in the root folder, one inside the `Reports` sub-folder

### Example 2: Search for CSV Files

**Scenario**: Find all CSV exports stored in a data drive

**Input**:
```
Drive Id:    b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3
File Format: .csv
```

**Output**:
```json
{
  "Files": [
    {
      "Drive ID": "b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3",
      "File ID": "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROX",
      "File Name": "Employee_Export.csv",
      "Web URL": "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/Exports/Employee_Export.csv"
    }
  ]
}
```

**Result**: One CSV file found recursively in the drive

### Example 3: No Files Found

**Scenario**: Drive exists but contains no files with the selected extension

**Input**:
```
Drive Id:    b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3
File Format: .tsv
```

**Output**:
```json
{
  "Files": []
}
```

**Result**: Empty list — no `.tsv` files found in the drive

### Example 4: Typical Two-Step Workflow

**Scenario**: Discover the Drive ID first, then search for files

**Step 1** — Call **Get All Shared Drives Information**:
```json
{
  "Drives": [
    {
      "Drive ID": "b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3",
      "Drive Name": "Finance Documents"
    }
  ]
}
```

**Step 2** — Call **Search Excel Files in Drive** using the Drive ID from Step 1:
```
Drive Id:    b!jxUeN0DgA02sUFccW1Pb1lXRHihDmDNCmAiRj77wj4mIw359lmkzQYFiwAJ5WZk3
File Format: .xlsx
```

**Result**: All `.xlsx` files from the Finance Documents drive returned

## Business Rules

1. **Single Format Per Request**: Only one file format can be selected per request — use separate calls for different formats
2. **Recursive Search**: All nested sub-folders are searched automatically — no folder path is needed
3. **Deduplication**: Duplicate files returned by the Graph API are automatically removed
4. **Permission-Based**: Only files the authenticated user has read access to are returned
5. **SharePoint Drives Supported**: Works with SharePoint document libraries (`b!...` Drive IDs) using the `/drives/{driveId}` Graph endpoint
6. **Pagination Handled**: Automatically follows `@odata.nextLink` cursors — all pages are retrieved
7. **Metadata Only**: File contents are not retrieved, only metadata (ID, name, URL)

## Limitations

1. **One Format at a Time**: Cannot search for multiple extensions in a single call — run separate requests per format if needed
2. **Single Drive**: Searches only within the specified drive — does not search across multiple drives
3. **No Folder Filtering**: Cannot restrict search to a specific sub-folder path
4. **No Name Filtering**: Cannot filter results by file name pattern within this request
5. **Graph Search Latency**: Newly uploaded files may not appear immediately in Graph search results
6. **Large Result Sets**: Drives with many files may have longer response times; pagination is handled automatically

## Best Practices

### 1. Get the Drive ID First
Always call **Get All Shared Drives Information** to obtain the correct Drive ID before calling this request.

### 2. Use File ID for All Subsequent Operations
Store the returned File ID and Drive ID together — both are required for most operations (Add Rows, Get Table Rows, etc.).

### 3. Handle Empty Results Gracefully
Always check if the Files list is empty before accessing file properties.

### 4. Search Per Format as Needed
If you need files of multiple formats, run one request per format and merge results in your workflow.

### 5. Cache Results When Appropriate
For drives that change infrequently, cache the returned file list to reduce repeated API calls.

## Common Use Cases

### 1. File Picker for Downstream Operations
```
Scenario: Let users select a specific Excel file from a shared drive
Action:   Call Search Excel Files in Drive (.xlsx), display File Names for selection
Result:   User picks a file; use its File ID and Drive ID in subsequent catalog requests
```

### 2. Automated Batch Processing
```
Scenario: Process every CSV file in a SharePoint library nightly
Action:   Search for .csv files, iterate the returned list, perform operations on each
Result:   All CSV files in the drive processed automatically, including files in sub-folders
```

### 3. File Existence Check
```
Scenario: Verify that a specific report file exists before starting a workflow
Action:   Search for .xlsx files, check if the target file name is in the returned list
Result:   Confirm presence of the file before proceeding
```

### 4. Cross-Format Inventory
```
Scenario: Audit all spreadsheet formats stored in a drive
Action:   Run four separate requests (.xlsx, .xls, .csv, .tsv) and combine results
Result:   Complete inventory of all spreadsheet files in the drive
```

## Related Catalog Requests

- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Get Drive IDs to use as input
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Alternative: lists files from all shared drives (no format filter, no deep search)
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Use returned File ID and Drive ID to add data
- [Get Table Rows](pages/GetTableRows.md) - Use returned File ID and Drive ID to read table data
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Use returned File ID and Drive ID to add sheet rows
- [Get Rows for Range](pages/GetRowsForRangeShared.md) - Use returned File ID and Drive ID to read a range

## Technical Implementation

### Class
- **Class**: `SharedArea`
- **Package**: `app.krista.extensions.essentials.spreadsheets.office365excel.catalog`
- **Method**: `searchExcelFilesInDrive(String driveId, String fileFormat)`
- **Service**: Delegates to `DriveService.searchExcelFilesInDrive(driveId, fileFormat)`

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /drives/{driveId}/root/search(q='{extension}')?$top=200`
- **Pagination**: Follows `@odata.nextLink` ($skiptoken) automatically
- **Permissions Required**: `Files.ReadWrite.All`

> **Note**: The `/drives/{driveId}` path is used (not `/me/drives/{driveId}`) to support SharePoint document libraries which are not part of the user's personal drive collection.

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **searchExcelFilesInDrive**: Logs search result count and any per-extension failures

## Troubleshooting

### Issue: Empty Files List Despite Files Existing

**Cause**: Drive ID is incorrect or the user lacks access

**Solution**:
1. Confirm the Drive ID using **Get All Shared Drives Information**
2. Verify the user has at least read access to the drive
3. Ensure the drive contains files with the selected extension

### Issue: `itemNotFound` (404) Error in Logs

**Cause**: The Drive ID does not exist or the user has no access to it

**Solution**:
1. Obtain a fresh Drive ID from **Get All Shared Drives Information**
2. Verify the SharePoint site is still active and accessible
3. Check that the user account has not lost access to the site

### Issue: `Insufficient Permissions` Error

**Cause**: `Files.ReadWrite.All` permission not granted

**Solution**:
1. Navigate to Microsoft Entra ID (Azure AD) → App registrations
2. Add `Files.ReadWrite.All` delegated permission
3. Grant admin consent
4. Re-authenticate the user

### Issue: Files in Sub-Folders Not Returned

**Cause**: Unlikely — the Graph search API searches recursively by design

**Solution**:
1. Confirm the sub-folder is inside the specified drive (not a different drive)
2. Verify the file has the correct extension matching the selected File Format
3. Wait a few minutes — newly uploaded files may not be indexed immediately

### Issue: Very Large Response / Slow Performance

**Cause**: Drive contains a large number of files

**Solution**:
1. This is expected — pagination is handled automatically
2. Consider caching results for drives that do not change frequently
3. If only a specific folder is needed, retrieve file IDs from **Get Shared Excel Files** and filter client-side

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) - Prerequisite: get Drive IDs
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Alternative file listing without format filter
