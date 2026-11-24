# Get All Excel Files For Name

## Overview

Retrieves all Excel files from the user's personal OneDrive that match a specified name or name pattern. This request is useful for finding specific files or discovering available workbooks.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

> **📝 Note**: Despite being marked as CHANGE_SYSTEM, this is a read-only query operation that does not modify any data.

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Name | Text | Yes | The name or partial name of the Excel file to search for | "Sales" |

### Parameter Details

#### Name
- **Format**: Full or partial file name
- **Case Sensitivity**: Case-insensitive search
- **Wildcards**: Not explicitly supported, but partial matches work
- **Extension**: Do not include file extension (.xlsx, .xls, etc.)
- **Empty String**: Returns all Excel files if empty
- **Example**: `"Sales"`, `"Report"`, `"Budget 2024"`

> **💡 Tip**: Use partial names to find multiple related files (e.g., "Sales" finds "Sales_Q1.xlsx", "Sales_Q2.xlsx", etc.)

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Files | List&lt;Entity(Excel File)&gt; | List of Excel files matching the search criteria |

### Excel File Entity Structure

Each Excel File entity contains the following properties:

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| File ID | Text | Unique identifier for the Excel file | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| File Name | Text | Name of the Excel file | "Sales_Report.xlsx" |
| Web URL | Text | Web URL to access the file in browser | "https://onedrive.live.com/..." |
| Created Date | Text | Date when the file was created | "2024-01-15T10:30:00Z" |
| Modified Date | Text | Date when the file was last modified | "2024-01-20T14:45:00Z" |

### Output Example

```json
{
  "Files": [
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCF",
      "File Name": "Sales_Report_Q1.xlsx",
      "Web URL": "https://onedrive.live.com/edit.aspx?resid=...",
      "Created Date": "2024-01-15T10:30:00Z",
      "Modified Date": "2024-01-20T14:45:00Z"
    },
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCD",
      "File Name": "Sales_Report_Q2.xlsx",
      "Web URL": "https://onedrive.live.com/edit.aspx?resid=...",
      "Created Date": "2024-04-10T09:00:00Z",
      "Modified Date": "2024-04-22T16:20:00Z"
    }
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| Name is null | "File name is required" | Provide a file name or partial name |
| User not authenticated | "Authentication required" | Ensure user is properly authenticated |
| Invalid access token | "Invalid or expired access token" | Re-authenticate to obtain new token |
| Insufficient permissions | "Insufficient permissions to access OneDrive" | Verify Files.ReadWrite permission is granted |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Name parameter is null (empty string is allowed)
- Invalid characters in file name

**Error Message**: `"File name is required"`

**Resolution**:
1. Provide a valid file name or partial name
2. Use empty string to search all files
3. Remove invalid characters from search term

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic or permission issues

**Common Scenarios**:
- No files match the search criteria
- User has no Excel files in OneDrive
- User account lacks Office 365 license
- OneDrive not provisioned for user

**Error Message**: `"No Excel files found matching '{name}'"` or `"OneDrive not accessible"`

**Resolution**:
1. Verify files exist with the specified name
2. Check user has Excel files in OneDrive
3. Ensure user has valid Office 365 license
4. Verify OneDrive is provisioned and accessible

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to retrieve Excel files"` or `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions (Files.ReadWrite not granted)
- User doesn't have access to OneDrive

**Error Message**: `"Unauthorized access"` or `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify Files.ReadWrite permission is granted
3. Ensure user has OneDrive access
4. Check user account is active and licensed

## Usage Examples

### Example 1: Find Files by Exact Name

**Scenario**: Find all files named "Budget"

**Input**:
```
Name: "Budget"
```

**Output**:
```json
{
  "Files": [
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCF",
      "File Name": "Budget.xlsx"
    },
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCD",
      "File Name": "Budget_2024.xlsx"
    }
  ]
}
```

**Result**: Two files containing "Budget" in their name

### Example 2: Find Files by Partial Name

**Scenario**: Find all sales-related files

**Input**:
```
Name: "Sales"
```

**Output**:
```json
{
  "Files": [
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCF",
      "File Name": "Sales_Report_Q1.xlsx"
    },
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCD",
      "File Name": "Sales_Data.xlsx"
    },
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCE",
      "File Name": "Monthly_Sales.xlsx"
    }
  ]
}
```

**Result**: Three files with "Sales" in their name

### Example 3: Get All Excel Files

**Scenario**: Retrieve all Excel files from OneDrive

**Input**:
```
Name: ""
```

**Output**:
```json
{
  "Files": [
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCF",
      "File Name": "Report.xlsx"
    },
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCD",
      "File Name": "Data.csv"
    },
    {
      "File ID": "01C8MOQGI3YRBDXAQERECWE3W4XCE",
      "File Name": "Budget.xlsx"
    }
  ]
}
```

**Result**: All Excel files in OneDrive returned

### Example 4: No Files Found

**Scenario**: Search for non-existent file

**Input**:
```
Name: "NonExistentFile"
```

**Output**:
```json
{
  "Files": []
}
```

**Result**: Empty list indicates no matching files

## Business Rules

1. **Partial Matching**: Searches for files containing the specified name
2. **Case-Insensitive**: Search is not case-sensitive
3. **Personal OneDrive Only**: Searches only in user's personal OneDrive
4. **All Formats**: Returns Excel files in all formats (.xlsx, .xls, .csv, .tsv)
5. **No Folder Filtering**: Searches across all folders in OneDrive
6. **Real-Time**: Results reflect current file availability
7. **Permission-Based**: Only files the user has access to are returned

## Limitations

1. **Personal OneDrive Only**: Does not search shared drives (use Shared Area requests)
2. **No Wildcards**: Does not support explicit wildcard characters (*, ?)
3. **No Regex**: Does not support regular expression patterns
4. **No Folder Specification**: Cannot limit search to specific folders
5. **No Date Filtering**: Cannot filter by creation or modification date
6. **Large Result Sets**: May return many files if search term is too broad
7. **No Sorting**: Results are not sorted in any specific order

## Best Practices

### 1. Use Specific Search Terms
Use specific names to narrow down results and improve performance.

### 2. Handle Empty Results
Always check if the Files list is empty before accessing file properties.

### 3. Cache File List
Store file list locally to avoid repeated API calls for the same search.

### 4. Use File ID for Operations
Always use File ID (not name) for subsequent operations, as names can change.

### 5. Validate File Existence
Before performing operations, verify the file still exists and is accessible.

### 6. Filter Client-Side
If you need advanced filtering, retrieve files and filter on the client side.

## Common Use Cases

### 1. File Selection Interface
```
Scenario: Display a dropdown of available Excel files for user selection
Action: Search for files and populate dropdown with File Names
Result: User can select a file to work with
```

### 2. Find Specific Report
```
Scenario: Find the latest sales report file
Action: Search for "Sales_Report" and select most recent by Modified Date
Result: Correct report file identified for processing
```

### 3. File Discovery
```
Scenario: Help users discover what Excel files they have
Action: Search with empty string to get all files
Result: Complete list of user's Excel files
```

### 4. Automated File Processing
```
Scenario: Process all budget files
Action: Search for "Budget" and iterate through results
Result: All budget files processed
```

## Related Catalog Requests

- [Get Excel File Information](pages/GetExcelFileInformation.md) - Get detailed file information
- [Get Cell Value](pages/GetCellValue.md) - Read cell values from files
- [Add Data to Cell](pages/AddDataToCell.md) - Write data to files
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Shared drive version
- [Get Sheets Information](pages/GetSheetsInformation.md) - Get sheet names from file

## Technical Implementation

### Helper Class
- **Class**: SheetArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog
- **Validation**: Input validation for file name
- **Service**: Delegates to WorkbookService.getWorkbooks()

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drive/root/search(q='{name}')`
- **Filter**: Filters for Excel file extensions
- **Permissions Required**: Files.ReadWrite

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **FILES_FOUND_COUNT**: Number of files matching search
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: Empty Files List

**Cause**: No files match the search criteria

**Solution**:
1. Verify files exist with the specified name
2. Try a broader search term (e.g., partial name)
3. Use empty string to see all Excel files
4. Check files are in personal OneDrive (not shared drives)

### Issue: "Insufficient Permissions" Error

**Cause**: Files.ReadWrite permission not granted

**Solution**:
1. Navigate to Microsoft Entra ID application
2. Add Files.ReadWrite delegated permission
3. Grant admin consent if required
4. Re-authenticate user

### Issue: Too Many Results

**Cause**: Search term is too broad

**Solution**:
1. Use more specific search terms
2. Filter results on client side
3. Implement pagination if needed

### Issue: Expected File Not Found

**Cause**: File name doesn't match or file is in shared drive

**Solution**:
1. Verify exact file name (case-insensitive)
2. Check file is in personal OneDrive
3. Use "Get Shared Excel Files" for shared drive files
4. Verify user has access to the file

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get Excel File Information](pages/GetExcelFileInformation.md) - Get detailed file info
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Search shared drives
- [Get Sheets Information](pages/GetSheetsInformation.md) - Get sheets from file

