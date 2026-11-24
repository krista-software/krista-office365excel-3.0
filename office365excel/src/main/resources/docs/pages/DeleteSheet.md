# Delete Sheet

## Overview

Deletes a specified worksheet from an Excel file stored in the user's personal OneDrive. This is a destructive operation that permanently removes the sheet and all its data.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes - Automatically retries on transient failures (rate limiting, temporary network issues)

> **⚠️ Warning**: This operation permanently deletes the sheet and cannot be undone. Ensure you have backups if needed.

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| File Id | Text | Yes | The ID of the Excel file containing the sheet to delete | "01C8MOQGI3YRBDXAQERECWE3W4XCF" |
| Sheet Name | Text | Yes | The name of the sheet to be deleted from the Excel file | "Sheet2" |

### Parameter Details

#### File Id
- **Format**: Unique identifier for the Excel file in OneDrive
- **How to obtain**: Use "Get All Excel Files For Name" catalog request
- **Validation**: Must be a valid Excel file ID accessible by the authenticated user
- **Example**: `01C8MOQGI3YRBDXAQERECWE3W4XCF`

#### Sheet Name
- **Format**: Name of the worksheet as it appears in Excel
- **Case Sensitivity**: Case-sensitive
- **Validation**: Sheet must exist in the Excel file
- **Special Characters**: Supported, but match exactly
- **Example**: `Sheet2`, `Old Data`, `Temp Sheet`

> **📝 Note**: Sheet name must match exactly, including case and special characters.

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Result | Text | Confirmation message of the delete operation |

### Output Example

```json
{
  "Result": "Sheet 'Sheet2' deleted successfully"
}
```

## Validation Rules

| Validation | Error Message | Resolution |
|------------|---------------|------------|
| File Id is empty or null | "File ID is required" | Provide a valid File ID |
| Sheet Name is empty or null | "Sheet Name is required" | Provide a valid sheet name |
| File not found | "File not found or inaccessible" | Verify File ID and user permissions |
| Sheet not found | "Sheet '{sheetName}' not found in the file" | Verify sheet name (case-sensitive) |
| Last sheet in workbook | "Cannot delete the last sheet in the workbook" | Excel requires at least one sheet |
| File is locked | "File is locked or in use" | Close file in Excel and retry |
| Insufficient permissions | "Insufficient permissions to modify file" | Verify user has write access |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters

**Common Scenarios**:
- Missing required parameters (File Id, Sheet Name)
- Empty or whitespace-only parameters
- Invalid File ID format

**Error Message**: `"File ID is required"`, `"Sheet Name is required"`

**Resolution**:
1. Verify all required parameters are provided
2. Ensure no parameters are empty or null
3. Check File ID is valid format

### Logic Errors (LOGIC_ERROR)

**Cause**: Business logic validation failures

**Common Scenarios**:
- Sheet doesn't exist in the specified file
- Attempting to delete the last sheet in workbook
- File not accessible or doesn't exist
- File is locked or checked out
- Insufficient permissions to modify file

**Error Message**: `"Sheet '{sheetName}' not found"`, `"Cannot delete the last sheet in the workbook"`

**Resolution**:
1. Verify sheet name is correct and case-sensitive
2. Ensure file has multiple sheets before deleting
3. Check file exists and is accessible
4. Verify file is not locked or in use
5. Ensure user has write permissions

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or system-level failures

**Common Scenarios**:
- Network connectivity issues
- Microsoft Graph API temporary unavailability
- Rate limiting (429 errors)
- Server errors (500, 503)
- Concurrent modification conflicts

**Error Message**: `"Failed to delete sheet"`, `"Microsoft Graph API error: {details}"`

**Resolution**:
1. Retry the operation (automatic retry is enabled)
2. Check network connectivity
3. Verify Microsoft Graph API status
4. Wait a few minutes if rate limited
5. Ensure no concurrent modifications

### Authorization Errors

**Cause**: Authentication or permission issues

**Common Scenarios**:
- Access token expired
- Insufficient permissions
- User doesn't have write access to the file
- File is read-only

**Error Message**: `"Unauthorized access"`, `"Insufficient permissions"`

**Resolution**:
1. Re-authenticate if token expired
2. Verify user has Files.ReadWrite permission
3. Check user has write access to the file
4. Ensure file is not read-only

## Usage Examples

### Example 1: Delete Temporary Sheet

**Scenario**: Remove a temporary worksheet no longer needed

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Temp Data"
```

**Output**:
```json
{
  "Result": "Sheet 'Temp Data' deleted successfully"
}
```

**Result**: Temporary sheet removed from workbook

### Example 2: Delete Old Data Sheet

**Scenario**: Remove outdated data sheet from report

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "2023 Data"
```

**Output**:
```json
{
  "Result": "Sheet '2023 Data' deleted successfully"
}
```

**Result**: Old data sheet removed, keeping workbook current

### Example 3: Attempt to Delete Last Sheet

**Scenario**: Try to delete the only sheet in workbook

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "Sheet1"
```

**Error**: LOGIC_ERROR - "Cannot delete the last sheet in the workbook"

**Result**: Operation fails because Excel requires at least one sheet

### Example 4: Delete Non-Existent Sheet

**Scenario**: Attempt to delete a sheet that doesn't exist

**Input**:
```
File Id: "01C8MOQGI3YRBDXAQERECWE3W4XCF"
Sheet Name: "NonExistent"
```

**Error**: LOGIC_ERROR - "Sheet 'NonExistent' not found in the file"

**Result**: Operation fails because sheet doesn't exist

## Business Rules

1. **Permanent Deletion**: Deleted sheets cannot be recovered
2. **Last Sheet Protection**: Cannot delete the last remaining sheet in a workbook
3. **Case-Sensitive**: Sheet name matching is case-sensitive
4. **Immediate Effect**: Deletion takes effect immediately
5. **No Undo**: There is no undo operation for sheet deletion
6. **References Broken**: Formulas referencing the deleted sheet will show errors
7. **Version History**: OneDrive version history may allow file-level recovery

## Limitations

1. **Cannot Delete Last Sheet**: Excel requires at least one sheet per workbook
2. **No Batch Delete**: Can only delete one sheet per request
3. **No Undo**: Deletion is permanent and cannot be undone
4. **Personal OneDrive Only**: This request is for personal OneDrive files only
5. **No Protected Sheets**: Cannot delete password-protected sheets
6. **No Confirmation**: Operation executes immediately without confirmation prompt

## Best Practices

### 1. Verify Sheet Exists First
Use "Get Sheets Information" to verify the sheet exists before attempting deletion.

### 2. Check Sheet Count
Ensure the workbook has multiple sheets before deleting to avoid errors.

### 3. Backup Important Data
Create a backup or version before deleting sheets with important data.

### 4. Handle Errors Gracefully
Implement error handling for cases where sheet doesn't exist or is the last sheet.

### 5. Confirm with Users
In user-facing applications, always confirm deletion before executing.

### 6. Update References
Update any formulas or references that point to the deleted sheet.

## Common Use Cases

### 1. Cleanup Temporary Sheets
```
Scenario: Remove temporary calculation or staging sheets
Action: Delete sheets named "Temp", "Staging", "Scratch"
Result: Workbook cleaned up, only final sheets remain
```

### 2. Remove Outdated Data
```
Scenario: Delete sheets containing old data after archiving
Action: Delete sheets for previous years or periods
Result: Workbook contains only current data
```

### 3. Workflow Automation
```
Scenario: Automated process creates and deletes working sheets
Action: Create temp sheet, process data, delete temp sheet
Result: Clean final workbook without intermediate sheets
```

### 4. Template Preparation
```
Scenario: Remove example sheets from template before distribution
Action: Delete "Example" or "Instructions" sheets
Result: Clean template ready for use
```

## Related Catalog Requests

- [Get Sheets Information](pages/GetSheetsInformation.md) - List all sheets in file
- [Create Sheet](pages/CreateSheet.md) - Create new sheets
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Find File ID
- [Get Excel File Information](pages/GetExcelFileInformation.md) - Get file details
- [Add Rows](pages/AddRows.md) - Add data to sheets

## Technical Implementation

### Helper Class
- **Class**: SheetOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service
- **Validation**: Input validation and sheet existence check
- **Service**: Delegates to WorksheetService for sheet operations

### Microsoft Graph API Endpoint
- **Endpoint**: `DELETE /me/drive/items/{fileId}/workbook/worksheets/{sheetName}`
- **Permissions Required**: Files.ReadWrite

### Telemetry Metrics
- **CATALOG_REQUEST_START**: Logged when request begins
- **CATALOG_REQUEST_SUCCESS**: Logged on successful completion
- **CATALOG_REQUEST_FAILURE**: Logged on error
- **SHEET_DELETED**: Logged when sheet is deleted
- **API_CALL_DURATION**: Tracks Microsoft Graph API call duration

## Troubleshooting

### Issue: "Sheet not found" Error

**Cause**: Sheet name doesn't match or sheet doesn't exist

**Solution**:
1. Verify sheet name is spelled correctly (case-sensitive)
2. Use "Get Sheets Information" to list available sheets
3. Check for extra spaces or special characters in name

### Issue: "Cannot delete the last sheet" Error

**Cause**: Attempting to delete the only sheet in workbook

**Solution**:
1. Create a new sheet before deleting the last one
2. Use "Get Sheets Information" to check sheet count
3. Consider deleting the entire file instead if appropriate

### Issue: "File is locked" Error

**Cause**: File is open in Excel or locked by another process

**Solution**:
1. Close the file in Excel or Excel Online
2. Wait a few minutes and retry
3. Check if file is checked out in SharePoint
4. Verify no other processes are accessing the file

### Issue: Formulas Show Errors After Deletion

**Cause**: Other sheets had formulas referencing the deleted sheet

**Solution**:
This is expected behavior. Update or remove formulas that referenced the deleted sheet.

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration
- [Authentication](pages/Authentication.md) - Authentication requirements
- [Get Sheets Information](pages/GetSheetsInformation.md) - List sheets before deleting
- [Create Sheet](pages/CreateSheet.md) - Create new sheets
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md) - Find files

