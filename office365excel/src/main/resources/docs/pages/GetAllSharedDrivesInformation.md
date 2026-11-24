# Get All Shared Drives Information

## Overview

Retrieves information about all shared drives accessible to the authenticated user, including OneDrive shared folders and SharePoint document libraries. This request provides essential metadata needed for subsequent operations on shared Excel files.

## Request Details

- **Area**: Shared
- **Type**: QUERY_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

This catalog request does not require any input parameters.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Drives | List&lt;Entity(Drive)&gt; | List of shared drives accessible to the user |

### Drive Entity Structure

Each Drive entity contains the following properties:

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| Drive ID | Text | Unique identifier for the drive | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| Drive Name | Text | Display name of the drive | "Marketing Team Drive" |
| Created By | Text | User who created or owns the drive | "John Doe" |
| Shared By | Text | User who shared the drive | "Jane Smith" |
| Drive Type | Text | Type of drive (e.g., documentLibrary, drive) | "documentLibrary" |
| Web URL | Text | Web URL to access the drive | "https://contoso.sharepoint.com/sites/Marketing" |

### Output Example

```json
{
  "Drives": [
    {
      "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
      "Drive Name": "Marketing Team",
      "Created By": "John Doe",
      "Shared By": "Jane Smith",
      "Drive Type": "documentLibrary",
      "Web URL": "https://contoso.sharepoint.com/sites/Marketing"
    },
    {
      "Drive ID": "b!xYzABm0EY0WK0z3E4xcAH5Ufuu38G9",
      "Drive Name": "Sales Documents",
      "Created By": "Alice Johnson",
      "Shared By": "Bob Williams",
      "Drive Type": "drive",
      "Web URL": "https://contoso-my.sharepoint.com/personal/alice_contoso_com"
    }
  ]
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| User not authenticated | "Authentication required" | Ensure user is properly authenticated |
| Invalid access token | "Invalid or expired access token" | Re-authenticate to obtain new token |
| Insufficient permissions | "Insufficient permissions to access shared drives" | Verify Files.ReadWrite.All permission is granted |

## Error Handling

### Input Errors (INPUT_ERROR)

This request has no input parameters, so input errors are not applicable.

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic or permission issues

**Common Scenarios**:
- User has no shared drives accessible
- User account lacks Office 365 license
- Conditional access policies blocking access

**Error Message**: `"No shared drives found"` or `"Access denied to shared drives"`

**Resolution**:
1. Verify user has access to at least one shared drive
2. Check user has valid Office 365 license
3. Ensure user account is active and not disabled
4. Review conditional access policies

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)

**Error Message**: `"Failed to retrieve shared drives"` or `"Microsoft Graph API error: {details}"`

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

### Example 1: List All Accessible Shared Drives

**Scenario**: Retrieve all shared drives to display in a selection list

**Input**: None

**Output**:
```json
{
  "Drives": [
    {
      "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
      "Drive Name": "HR Documents",
      "Created By": "HR Admin",
      "Shared By": "IT Department",
      "Drive Type": "documentLibrary",
      "Web URL": "https://contoso.sharepoint.com/sites/HR"
    },
    {
      "Drive ID": "b!aB1cDm0EY0WK0z3E4xcAH5Ufuu38H2",
      "Drive Name": "Finance Reports",
      "Created By": "Finance Team",
      "Shared By": "CFO",
      "Drive Type": "documentLibrary",
      "Web URL": "https://contoso.sharepoint.com/sites/Finance"
    }
  ]
}
```

**Result**: Two shared drives retrieved successfully

### Example 2: Find Drive ID for Subsequent Operations

**Scenario**: Get Drive ID to use in "Get Shared Excel Files" request

**Input**: None

**Output**:
```json
{
  "Drives": [
    {
      "Drive ID": "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8",
      "Drive Name": "Project Files",
      "Created By": "Project Manager",
      "Shared By": "Team Lead"
    }
  ]
}
```

**Next Step**: Use Drive ID `b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8` in subsequent catalog requests

**Result**: Drive ID obtained for use in other operations

### Example 3: No Shared Drives Available

**Scenario**: User has no shared drives accessible

**Input**: None

**Output**:
```json
{
  "Drives": []
}
```

**Result**: Empty list indicates no shared drives are accessible to the user

## Business Rules

1. **Permission-Based**: Only drives the authenticated user has access to are returned
2. **Real-Time Data**: Results reflect current permissions and sharing status
3. **Includes All Types**: Returns both OneDrive shared folders and SharePoint document libraries
4. **No Filtering**: All accessible shared drives are returned (no built-in filtering)
5. **Metadata Only**: Only drive metadata is returned, not file contents
6. **Created By Field**: May be empty if creator information is not available

## Limitations

1. **Personal OneDrive**: User's personal OneDrive is not included in shared drives list
2. **Deleted Drives**: Recently deleted drives may still appear temporarily
3. **Permission Changes**: Recent permission changes may take a few minutes to reflect
4. **Large Organizations**: In organizations with many shared drives, response may be paginated
5. **External Sharing**: Externally shared drives may have limited metadata
6. **Guest Users**: Guest users may see fewer drives than full organization members

## Best Practices

### 1. Cache Drive Information
Store drive information locally to avoid repeated API calls. Refresh periodically or when needed.

### 2. Handle Empty Results
Always check if the Drives list is empty before attempting to access drive properties.

### 3. Use Drive ID for Operations
Always use the Drive ID (not the name) for subsequent operations, as names can change.

### 4. Display User-Friendly Names
Show Drive Name to users for selection, but use Drive ID internally for operations.

### 5. Implement Error Handling
Handle scenarios where no drives are accessible or permissions have changed.

### 6. Refresh Periodically
Shared drive access can change; refresh the list periodically in long-running applications.

## Common Use Cases

### 1. Drive Selection Interface
```
Scenario: Display a dropdown list of available shared drives for user selection
Action: Call Get All Shared Drives Information and populate dropdown with Drive Names
Result: User can select a drive to work with
```

### 2. Automated File Processing
```
Scenario: Process Excel files across all accessible shared drives
Action: Retrieve all drives, then iterate through each to get Excel files
Result: Comprehensive file processing across organization
```

### 3. Access Verification
```
Scenario: Verify user has access to a specific shared drive before proceeding
Action: Retrieve all drives and check if target Drive ID is in the list
Result: Confirm access before attempting operations
```

### 4. Drive Discovery
```
Scenario: Help users discover what shared drives they have access to
Action: Display all accessible drives with metadata
Result: Users can see all available shared resources
```

## Related Catalog Requests

- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Get Excel files from shared drives
- [Add Rows to Table](pages/AddRowsToTableShared.md) - Add data to tables in shared files
- [Get Table Rows](pages/GetTableRows.md) - Retrieve data from tables in shared files
- [Add Rows to Sheet](pages/AddRowsToSheet.md) - Add data to sheets in shared files

## Technical Implementation

### Helper Class
- **Class**: SharedArea
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog
- **Validation**: Minimal validation (authentication check only)
- **Service**: Delegates to DriveService.getAllSharedDrivesInformation()

### Microsoft Graph API Endpoint
- **Endpoint**: `GET /me/drives`
- **Filter**: Filters for shared drives only
- **Permissions Required**: Files.ReadWrite.All

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration
- **DRIVES_RETRIEVED_COUNT**: Number of drives returned

## Troubleshooting

### Issue: Empty Drives List

**Cause**: User has no shared drives accessible

**Solution**:
1. Verify user has been granted access to at least one shared drive
2. Check SharePoint sharing settings in organization
3. Ensure user is not a guest with restricted access
4. Verify user account is active and licensed

### Issue: "Insufficient Permissions" Error

**Cause**: Files.ReadWrite.All permission not granted

**Solution**:
1. Navigate to Microsoft Entra ID application
2. Add Files.ReadWrite.All delegated permission
3. Grant admin consent
4. Re-authenticate user

### Issue: Missing Drive Information

**Cause**: Some drives may have limited metadata

**Solution**:
1. This is expected behavior for certain drive types
2. Use available fields (Drive ID is always present)
3. Handle missing fields gracefully in your code

### Issue: Outdated Drive List

**Cause**: Cached results or recent permission changes

**Solution**:
1. Clear any local caching
2. Re-run the request to get fresh data
3. Wait a few minutes for permission changes to propagate

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) - Next step after getting drives
- [Get Excel File Information](pages/GetExcelFileInformation.md) - Get file details from personal OneDrive

