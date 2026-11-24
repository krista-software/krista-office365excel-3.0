# Troubleshooting Guide

This guide helps you resolve common issues when using the Office365Excel extension.

## Authentication Issues

### "Your authentication session has expired"
**Error Message**: "Your authentication session has expired. Please re-authorize the connection or contact your system administrator if this problem persists."

**Cause**: The refresh token has expired or been revoked.

**Solution**:
1. Go to the extension configuration
2. Click "Validate Attributes" to re-authenticate
3. Complete the Microsoft login process
4. Save the configuration
5. If the issue persists after re-authorization, contact your system administrator

### "Authorization required. Please click 'Validate Attributes'"
**Error Message**: "Authorization required. Please click 'Validate Attributes' to authenticate before saving your configuration."

**Cause**: The extension needs initial authorization or re-authorization.

**Solution**:
1. In the extension configuration, click "Validate Attributes"
2. Complete the Microsoft login process
3. Grant the requested permissions
4. Save the configuration

### "Your Microsoft account has been disabled or locked"
**Error Message**: "Your Microsoft account has been disabled or locked. Please contact your system administrator to unlock your account or verify if your password needs to be reset."

**Cause**: Your Microsoft account has been disabled by an administrator, or your password has expired.

**Solution**:
1. Contact your system administrator to unlock your account
2. Verify if your password needs to be reset
3. Check if your account has been temporarily locked due to failed login attempts
4. Once unlocked, re-authorize the connection

### "The Microsoft application registration is missing or misconfigured"
**Error Message**: "The Microsoft application registration is missing or misconfigured. Please contact your system administrator to verify the application setup in Azure Active Directory."

**Cause**: The Azure application registration is not properly configured or has been deleted.

**Solution**:
1. Contact your system administrator
2. Verify the application exists in Azure Active Directory
3. Ensure the Client ID and Tenant ID in the configuration are correct
4. Confirm the application has the required API permissions
5. Check that the Client Secret hasn't expired

### "Your Microsoft account has been deleted or deactivated"
**Error Message**: "Your Microsoft account has been deleted or deactivated. Please contact your system administrator to restore access or use a different account."

**Cause**: The Microsoft account no longer exists in the organization.

**Solution**:
1. Contact your system administrator to verify account status
2. If the account was deleted, request restoration or use a different account
3. Update the extension configuration with the new account credentials

### "Your access permissions have been revoked"
**Error Message**: "Your access permissions have been revoked or the application consent has been withdrawn. Please contact your system administrator to restore the necessary permissions."

**Cause**: The application consent has been withdrawn or permissions have been revoked.

**Solution**:
1. Contact your system administrator
2. Request restoration of the necessary permissions
3. Re-grant consent to the application
4. Re-authorize the connection in the extension configuration

### "Your Microsoft organization could not be found"
**Error Message**: "Your Microsoft organization could not be found. Please verify the Tenant ID in your configuration or contact your system administrator."

**Cause**: The Tenant ID is incorrect or the organization doesn't exist.

**Solution**:
1. Verify the Tenant ID in the extension configuration
2. Check for typos in the Tenant ID
3. Contact your system administrator to confirm the correct Tenant ID
4. Update the configuration with the correct Tenant ID

### "Microsoft services are temporarily unavailable"
**Error Message**: "Microsoft services are temporarily unavailable. Please wait a few minutes and try again. If the issue persists, contact your system administrator."

**Cause**: Microsoft Graph API or Microsoft 365 services are experiencing issues.

**Solution**:
1. Wait 5-10 minutes and try again
2. Check Microsoft 365 service status at https://status.office.com
3. Verify your internet connection
4. If the issue persists for more than 30 minutes, contact your system administrator

### "The application credentials are invalid or expired"
**Error Message**: "The application credentials are invalid or expired. Please contact your system administrator to update the Client Secret in Azure Active Directory."

**Cause**: The Client Secret has expired or is incorrect.

**Solution**:
1. Contact your system administrator
2. Generate a new Client Secret in Azure Active Directory
3. Update the extension configuration with the new Client Secret
4. Save and re-authorize the connection

### "Your Microsoft password was recently changed"
**Error Message**: "Your Microsoft password was recently changed. Please re-authorize the connection to continue."

**Cause**: Your Microsoft password was changed, invalidating the current authentication.

**Solution**:
1. Go to the extension configuration
2. Click "Validate Attributes" to re-authorize
3. Log in with your new password
4. Save the configuration

## File Access Issues

### "You are not authorized to access this Excel file"
**Error Message**: "You are not authorized to access this Excel file. Please verify that you have the necessary permissions or contact your system administrator."

**Cause**: You don't have permission to access the file, or the file doesn't exist.

**Solution**:
1. Verify the file ID is correct
2. Check that the file exists in OneDrive or SharePoint
3. Ensure you have read/write permissions to the file
4. Request access from the file owner
5. Use "Get Excel File Information" to find the correct file ID
6. Contact your system administrator if permissions are managed centrally

### "Excel workbook not found at the specified path"
**Cause**: The file path is incorrect, the file has been moved/deleted, or you don't have access.

**Solution**:
1. Check the file path for typos (paths are case-sensitive)
2. Verify the file still exists at that location
3. Ensure the file hasn't been moved to a different folder
4. Check that you have permission to access the folder
5. Update the path if the file has been moved
6. Use "Get All Excel Files For Name" to search for the file

## Sheet and Table Issues

### "Sheet 'SheetName' not found in this workbook"
**Cause**: The sheet name is incorrect or the sheet doesn't exist.
**Solution**:
1. Check the sheet name for typos (case-sensitive)
2. Remove any leading/trailing spaces
3. Use "Get Sheets Information" to see all available sheets
4. Verify the sheet hasn't been deleted or renamed

### "Table 'TableName' not found or inaccessible"
**Cause**: The table name is incorrect or you don't have access.
**Solution**:
1. Verify the table name is spelled correctly (case-sensitive)
2. Check that the table exists in the workbook
3. Ensure you have permission to access the table

### "Column 'ColumnName' not found in the table"
**Cause**: The column name is incorrect or doesn't exist in the table.
**Solution**:
1. Check the column name spelling and capitalization
2. Verify the column exists in the specified table
3. Remove any extra spaces from the column name

## Data Format Issues

### "Invalid cell reference"
**Error Message**: "Invalid cell reference: [cell_id]. Please use valid Excel cell format (e.g., A1, B2, AA10)."

**Cause**: The cell reference format is incorrect.

**Solution**: Use proper Excel cell format:
- Single cell: A1, B2, AA10, XFD1048576
- Must have column letter(s) followed by row number
- Column letters: A-Z, AA-AZ, BA-BZ, ..., XFA-XFD
- Row numbers: 1-1048576
- Examples of valid cells: A1, Z99, AA100, XFD1048576
- Examples of invalid cells: 1A, A, 123, A1048577

### "Invalid Excel range format"
**Error Message**: "Invalid Excel range format: [range]. Please use valid Excel range format (e.g., A1:B10, A:A, 1:1)."

**Cause**: The range format is not valid.

**Solution**: Use valid Excel range formats:
- Cell range: A1:B10, C5:D20
- Single cell: A1 (also valid as a range)
- Entire column: A:A, B:Z
- Entire row: 1:1, 5:10
- Mixed ranges are not supported (e.g., A:B10 is invalid)
- Ensure start cell comes before end cell (A1:B10, not B10:A1)

### "Invalid column range format"
**Error Message**: "Invalid column range format: [range]. Please use valid column format (e.g., A, B, AA)."

**Cause**: The column reference is not in the correct format.

**Solution**:
- Use column letters only: A, B, C, ..., Z, AA, AB, ..., XFD
- For column ranges, use colon separator: A:C, B:Z
- Do not include row numbers in column references
- Valid examples: A, AA, XFD, A:C
- Invalid examples: A1, 1, A:C5

### "No row data provided"
**Error Message**: "No row data provided. Please provide either 'Rows' or 'Rows from CSV' to add data to the sheet."

**Cause**: Neither 'Rows' nor 'Rows from CSV' parameters were provided.

**Solution**: Provide data using either:
- The 'Rows' parameter with comma-separated text data
- The 'Rows from CSV' parameter with CSV files
- At least one of these parameters must be provided
- Both parameters can be used together (CSV data will be processed first)

### "One or more input parameters are invalid"
**Error Message**: "One or more input parameters are invalid. Please check that all required fields are provided with correct values and try again."

**Cause**: Required parameters are missing or have invalid values.

**Solution**:
1. Review the catalog request documentation for required parameters
2. Ensure all required fields are provided
3. Verify parameter values match the expected format
4. Check for typos in parameter names
5. Ensure numeric values are within valid ranges
6. Verify file IDs, sheet names, and table names are correct

### "One or more fields contain invalid data"
**Error Message**: "One or more fields contain invalid data. Please verify your input values match the expected format and try again."

**Cause**: Field values don't match the expected data type or format.

**Solution**:
1. Check that numeric fields contain only numbers
2. Verify date fields are in the correct format
3. Ensure text fields don't exceed maximum length
4. Remove special characters if not allowed
5. Check that cell references and ranges are valid
6. Verify file paths and names are correct

## Formula Issues

### "Unable to execute Excel formula"
**Error Message**: "Unable to execute Excel formula: [error_details]. Please check that your formula syntax is correct, all referenced cells exist, and you have permission to modify the workbook."

**Cause**: The formula contains syntax errors, references invalid cells, or you lack permissions.

**Solution**:
1. Ensure the formula starts with '=' (e.g., =SUM(A1:A10))
2. Verify all cell references exist in the worksheet
3. Check that function names are spelled correctly
4. Ensure parentheses are balanced
5. Keep formulas under 8,192 characters
6. Test the formula manually in Excel first
7. Verify you have permission to modify the workbook
8. Check that referenced sheets exist if using cross-sheet formulas
9. Ensure circular references are avoided

**Common Formula Errors**:
- `#NAME?` - Function name is misspelled or doesn't exist
- `#REF!` - Cell reference is invalid or deleted
- `#VALUE!` - Wrong data type used in formula
- `#DIV/0!` - Division by zero
- `#N/A` - Value not available to formula
- `#NUM!` - Invalid numeric value

### "An unexpected error occurred while executing the formula"
**Error Message**: "An unexpected error occurred while executing the formula: [error_details]. Please check your formula syntax and try again. If the issue continues, contact your system administrator."

**Cause**: Unexpected error during formula execution.

**Solution**:
1. Simplify the formula and test with basic operations first
2. Break complex formulas into smaller parts
3. Verify all cell references and ranges are valid
4. Check for special characters that might cause issues
5. Try the formula in Excel to ensure it works
6. Contact your system administrator if the issue persists

## Network and Service Issues

### "Unable to process CSV files"
**Cause**: CSV files are corrupted, inaccessible, in wrong format, or too large.

**Solution**:
1. Verify CSV files are properly formatted with consistent delimiters
2. Check that files are accessible and not locked by another application
3. Ensure files are not corrupted (try opening in a text editor)
4. Verify the file encoding is UTF-8 or ASCII
5. Check that column counts are consistent across all rows
6. Try with a smaller CSV file first to isolate the issue
7. Ensure CSV files don't exceed size limits
8. Remove any special characters or formatting that might cause issues

### "Unable to determine file type"
**Error Message**: "Unable to determine file type: The file '[filename]' does not have a file extension. Please ensure the file has a valid extension (e.g., .xlsx, .csv)."

**Cause**: The file doesn't have a file extension.

**Solution**:
1. Ensure the file has a proper extension (.xlsx for Excel files)
2. Rename the file to include the extension
3. Verify the file is actually an Excel file
4. Check that the file isn't corrupted

## Performance Issues

### Slow Response Times
**Cause**: Large datasets or network latency.
**Solution**:
1. Use batch operations for large datasets
2. Consider breaking large operations into smaller chunks
3. Use async operations when available
4. Check your network connection

### Timeout Errors
**Cause**: Operations taking too long to complete.
**Solution**:
1. Reduce the amount of data being processed
2. Use async operations for large datasets
3. Check network connectivity
4. Try the operation during off-peak hours

## Getting Additional Help

If you continue to experience issues:

1. **Check the logs**: Look for detailed error messages in the application logs
2. **Verify permissions**: Ensure you have the necessary permissions for all operations
3. **Test with simple data**: Try the operation with minimal data first
4. **Contact support**: Reach out to your system administrator with specific error messages
5. **Check documentation**: Review the specific catalog request documentation for detailed requirements

## Common Error Codes

| Error Code  | Description           | Solution                                    |
|-------------|-----------------------|---------------------------------------------|
| 401         | Unauthorized          | Re-authenticate the connection              |
| 403         | Forbidden             | Check permissions with administrator        |
| 404         | Not Found             | Verify file/sheet/table names and IDs       |
| 429         | Too Many Requests     | Wait and retry, or reduce request frequency |
| 500         | Internal Server Error | Contact administrator, may be temporary     |
| 503         | Service Unavailable   | Microsoft services temporarily down         |

## Azure AD Error Codes (AADSTS)

| Error Code    | Error Message | Meaning |
|---------------|---------------|---------|
| AADSTS50020   | User account deleted | The Microsoft account has been deleted from the directory |
| AADSTS50053   | Account locked | The account has been locked due to too many failed sign-in attempts |
| AADSTS50055   | Password expired | The user's password has expired and needs to be reset |
| AADSTS50057   | Account disabled | The user account has been disabled by an administrator |
| AADSTS50173   | Password changed | The user's password was recently changed |
| AADSTS65001   | Consent revoked | The user or administrator has revoked consent for the application |
| AADSTS70019   | Consent required | The application requires user consent |
| AADSTS90002   | Tenant not found | The specified tenant ID doesn't exist |
| AADSTS90094   | Role not found | The required role assignment is missing |
| AADSTS700016  | App not found | The application registration doesn't exist in the directory |
| AADSTS700082  | Refresh token expired | The refresh token has expired |
| AADSTS700084  | Refresh token revoked | The refresh token has been revoked |
| AADSTS7000215 | Invalid client secret | The client secret is invalid or has expired |

## Quick Reference: Error Message Patterns

### Authentication Errors
- **"Your authentication session has expired"** → Re-authorize the connection
- **"Authorization required"** → Click 'Validate Attributes' to authenticate
- **"Your Microsoft account has been disabled"** → Contact administrator to unlock account
- **"The Microsoft application registration is missing"** → Verify Azure AD app setup
- **"Your access permissions have been revoked"** → Request permission restoration
- **"Your Microsoft organization could not be found"** → Verify Tenant ID
- **"The application credentials are invalid"** → Update Client Secret in Azure AD
- **"Your Microsoft password was recently changed"** → Re-authorize with new password

### Data Validation Errors
- **"Invalid cell reference"** → Use format like A1, B2, AA10
- **"Invalid Excel range format"** → Use format like A1:B10, A:A, 1:1
- **"Invalid column range format"** → Use column letters only (A, B, AA)
- **"No row data provided"** → Provide either 'Rows' or 'Rows from CSV'
- **"One or more input parameters are invalid"** → Check required fields and formats
- **"One or more fields contain invalid data"** → Verify data types and formats

### File and Resource Errors
- **"You are not authorized to access this Excel file"** → Request file permissions
- **"Excel workbook not found"** → Verify file path and existence
- **"Sheet not found"** → Check sheet name spelling (case-sensitive)
- **"Table not found"** → Verify table name and existence
- **"Column not found"** → Check column name in table
- **"Unable to determine file type"** → Ensure file has .xlsx extension

### Formula Errors
- **"Unable to execute Excel formula"** → Check formula syntax and cell references
- **"An unexpected error occurred while executing the formula"** → Simplify formula and test in Excel

### Service Errors
- **"Microsoft services are temporarily unavailable"** → Wait and retry
- **"Unable to process CSV files"** → Check CSV format and encoding

## Best Practices for Error Resolution

1. **Read the complete error message** - Error messages include specific details about what went wrong
2. **Check the obvious first** - Typos, case sensitivity, missing required fields
3. **Verify permissions** - Ensure you have the necessary access rights
4. **Test with simple data** - Start with minimal data to isolate the issue
5. **Check service status** - Visit https://status.office.com for Microsoft 365 status
6. **Review recent changes** - Consider what changed since it last worked
7. **Consult documentation** - Review the specific catalog request documentation
8. **Contact support with details** - Provide the complete error message and steps to reproduce

Remember: Error messages are designed to be user-friendly and actionable. Read them carefully as they often contain specific guidance for resolving the issue.
