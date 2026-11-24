# Get Rows For Range By Path

## Overview

Retrieves all rows of data within a specified cell range from an Excel worksheet, locating the file by its
SharePoint or OneDrive path instead of a File ID. This removes the need to look up the File ID separately —
provide the path to the file and the extension resolves the rest.

URL-encoded paths (e.g. `Shared%20Documents`) are automatically decoded, so plain spaces in path segments work too.

## Request Details

- **Area**: Sheet
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes — Automatically retries on transient failures (rate limiting, temporary network issues)

## Input Parameters

| Parameter Name  | Type     | Required | Description                                               | Example                                                |
|-----------------|----------|----------|-----------------------------------------------------------|--------------------------------------------------------|
| Source Path     | Text     | Yes      | Full path to the Excel file                               | `/Engineering/Shared Documents/Reports/Q1_Report.xlsx` |
| Sheet Name      | Text     | Yes      | The name of the worksheet containing the range            | `Sheet1`                                               |
| Range           | Text     | Yes      | The Excel range to retrieve rows from                     | `A1:D10`                                               |
| File Search | Pick One | Yes      | Where to look for the file: `My Drive` or `Site Files`   | `Site Files`                                           |
| Allow Retry     | Boolean  | No       | Whether to re-enter on failure                            | `true`                                                 |

### Parameter Details

#### Source Path

The full path to the Excel file. The required format depends on the **File Search** selected.

| Strategy   | Path Format                                           | Example                                                 |
|------------|-------------------------------------------------------|---------------------------------------------------------|
| My Drive   | `/folder/subfolder/file.xlsx`                         | `/Documents/Reports/Q1_Report.xlsx`                     |
| Site Files | `/SiteName/LibraryName/optional/subfolders/file.xlsx` | `/Engineering/Shared Documents/Reports/Q1_Report.xlsx`  |

- **Leading slash**: Optional — the extension normalises it automatically.
- **URL-encoding**: `%20` and other percent-encoded characters are decoded before lookup.
- **Case sensitivity**: Folder and file names are matched case-insensitively.

#### Sheet Name

- **Format**: Name of the worksheet exactly as it appears in Excel.
- **Case Sensitivity**: Case-sensitive.
- **Validation**: Sheet must exist in the Excel file.
- **Example**: `Sheet1`, `Sales Data`, `Q1 Report`

#### Range

- **Format**: Excel range notation (e.g., `A1:D10`, `C5:F20`).
- **Pattern**: `[Column][Row]:[Column][Row]`
- **Single Cell**: Supported (e.g., `A1`).
- **Validation**: Must be a valid Excel range within the sheet.
- **Example**: `A1:D10`, `C5:F20`, `A1:Z100`

> **Note**: Range format is case-insensitive (`A1:D10` = `a1:d10`)

#### File Search

Determines which Microsoft Graph API endpoint is used to locate the file.

| Value          | When to Use                                                                                                                       |
|----------------|-----------------------------------------------------------------------------------------------------------------------------------|
| **My Drive**   | The file is in your own personal OneDrive. Path is relative to your OneDrive root.                                                |
| **Site Files** | The file lives in a SharePoint site document library. The first path segment is the site name, the second is the library name.    |

> **Tip**: Use **My Drive** for files you own in your personal OneDrive. Use **Site Files** for files in SharePoint
> site libraries or Microsoft Teams channel files.

## Output Parameters

| Parameter Name | Type                  | Description                                              |
|----------------|-----------------------|----------------------------------------------------------|
| Rows           | List&lt;Composite&gt; | Array of rows, where each row is an array of cell values |

### Output Structure

The output is a two-dimensional array:

- Each element in the outer array represents a row.
- Each element in the inner array represents a cell value in that row.

### Output Example

```json
{
  "Rows": [
    ["Name", "Department", "Salary", "Location"],
    ["John Doe", "Engineering", "75000", "New York"],
    ["Jane Smith", "Marketing", "65000", "San Francisco"],
    ["Bob Johnson", "Sales", "70000", "Chicago"]
  ]
}
```

## Validation Rules

| Validation                     | Error Message                                                            | Resolution                                        |
|--------------------------------|--------------------------------------------------------------------------|---------------------------------------------------|
| Source Path is empty or null   | "Source path must not be blank"                                          | Provide a valid path to the Excel file            |
| Source Path has no filename    | "Source path must include a filename"                                    | Ensure the path ends with a filename              |
| Sheet Name is empty or null    | "Sheet Name is required"                                                 | Provide the worksheet name                        |
| Range is empty or null         | "Range is required"                                                      | Provide a valid Excel range                       |
| File Search not selected   | "File Search is required. Choose one of: My Drive, Site Files"       | Select My Drive or Site Files                     |
| Unknown File Search value  | "Unknown File Search '...'. Choose one of: My Drive, Site Files"     | Use one of the two supported strategy values      |
| File not found (My Drive)      | "File not found in My Drive at path: ..."                                | Verify the path and that the file is in your OneDrive |
| Site Files — path too short    | "Site Files lookup requires at least two path segments"                  | Include site name and library name in the path    |
| File not found (Site Files)    | "File not found in Site Files at path: ..."                              | Verify site name, library name, and file path     |
| Invalid range format           | "Invalid range format"                                                   | Use valid Excel range notation (e.g., A1:D10)     |

## Error Handling

### Input Errors (INPUT_ERROR)

**Cause**: Invalid or missing input parameters.

**Common Scenarios**:

- Source Path is blank or missing a filename.
- Range format is invalid (e.g., `1A:10D` instead of `A1:D10`).
- File Search not selected.

**Resolution**:

1. Verify the Source Path ends with the filename (e.g., `/file.xlsx`).
2. Check the Range follows Excel notation (column letter + row number).
3. Select a File Search.

### Logic Errors (LOGIC_ERROR)

**Cause**: The file, site, library, or sheet could not be found.

**Common Scenarios**:

- File does not exist at the given path.
- Site name typo (Site Files strategy).
- SharePoint library name mismatch — e.g. URL path uses `Shared Documents` but the display name differs.
- Sheet name is incorrect or misspelled.
- Range is outside the sheet bounds.

**Resolution**:

1. Double-check the Source Path spelling, including site and library names.
2. For **Site Files**, verify the site name matches the SharePoint site display name exactly.
3. For **My Drive**, ensure the file exists in your personal OneDrive at the specified path.
4. Verify the Sheet Name is correct and case-sensitive.

### System Errors (SYSTEM_ERROR)

**Cause**: Microsoft Graph API or network-level failures.

**Common Scenarios**:

- Network connectivity issues.
- Microsoft Graph API temporary unavailability.
- Rate limiting (HTTP 429).
- Large range causing timeout.

**Resolution**:

1. The request retries automatically on transient failures.
2. Check network connectivity.
3. Reduce the range size if timeouts occur.

### Authorization Errors

**Cause**: The authenticated user does not have access to the file or site.

**Common Scenarios**:

- Access token expired.
- User is not a member of the SharePoint site (Site Files strategy).
- File permissions changed after the workflow was created.

**Resolution**:

1. Re-authenticate if the token has expired.
2. Verify the user has read access to the file in SharePoint or OneDrive.
3. For Site Files, confirm the user is a member of the SharePoint site.

## Usage Examples

### Example 1: Site Files — SharePoint Document Library

**Scenario**: Read quarterly report data from a SharePoint site.

**Input**:

```
Source Path:     /Engineering/Shared Documents/Reports/Q1_Report.xlsx
Sheet Name:      Summary
Range:           A1:E10
File Search: Site Files
```

**Notes**:
- `Engineering` is the SharePoint site display name.
- `Shared Documents` is the document library name (URL path segment).
- `/Reports/Q1_Report.xlsx` is the file path inside the library.

**Output**:

```json
{
  "Rows": [
    ["Quarter", "Revenue", "Expenses", "Profit", "Growth"],
    ["Q1 2026", "120000", "85000", "35000", "12%"],
    ["Q4 2025", "110000", "80000", "30000", "8%"]
  ]
}
```

---

### Example 2: My Drive — Personal OneDrive

**Scenario**: Read budget data from a personal OneDrive file.

**Input**:

```
Source Path:     /Finance/Budget 2026.xlsx
Sheet Name:      Sheet1
Range:           A1:C5
File Search: My Drive
```

**Notes**:
- No site name or library name needed — path is relative to your OneDrive root.

**Output**:

```json
{
  "Rows": [
    ["Category", "Planned", "Actual"],
    ["Salaries", "500000", "495000"],
    ["Marketing", "50000", "48000"],
    ["Operations", "80000", "76000"]
  ]
}
```

---

### Example 3: My Drive — File at Root

**Scenario**: Read data from a file stored directly in the OneDrive root (no subfolder).

**Input**:

```
Source Path:     /Tracker.xlsx
Sheet Name:      Dashboard
Range:           B2:D10
File Search: My Drive
```

**Notes**:
- If the file is at the root of your OneDrive with no subfolder, just provide the filename with a leading slash.

---

### Example 4: Site Files — File in Subfolder

**Scenario**: File is nested inside a subfolder within the SharePoint library.

**Input**:

```
Source Path:     /Sales Team/Documents/2026/Q1/Pipeline.xlsx
Sheet Name:      Active Deals
Range:           A1:F50
File Search: Site Files
```

**Notes**:
- `Sales Team` is the SharePoint site name.
- `Documents` is the library.
- `2026/Q1/` is the subfolder path within the library.

## Business Rules

1. **Path Resolution First**: The file is resolved by path before any sheet or range operations are performed.
2. **Includes All Data**: Returns all cell values in the specified range.
3. **Preserves Structure**: Maintains row and column order.
4. **Empty Cells**: Returned as empty strings.
5. **Formula Values**: Returns calculated values, not formula text.
6. **Data Types**: All values are returned as strings.
7. **No Formatting**: Cell formatting, colors, and styles are not included.
8. **Case-Insensitive Path**: Folder and file name matching is case-insensitive.
9. **URL Decoding**: Percent-encoded characters (e.g. `%20`) in the Source Path are decoded automatically.

## Choosing the Right File Search

| Scenario                                                     | Strategy    |
|--------------------------------------------------------------|-------------|
| File is in your personal OneDrive                            | My Drive    |
| You uploaded the file yourself to OneDrive                   | My Drive    |
| File is in a SharePoint site or Microsoft Teams channel      | Site Files  |
| File is in a team's SharePoint library (General / Documents) | Site Files  |

## File Size Limits

| Size Range      | Status        | Notes                                                        |
|-----------------|---------------|--------------------------------------------------------------|
| Up to 100 MB    | Supported     | Fully supported                                              |
| 100 MB – 150 MB | Best Effort   | May succeed but not guaranteed; prone to timeout             |
| Above 150 MB    | Not Supported | Expected to fail due to Microsoft Graph API timeout limits   |

> For files above 100 MB, consider splitting the workbook or reducing data volume.

## Limitations

1. **Site Files path depth**: The first segment must be the SharePoint site name; the second must be the document library name.
2. **Large Ranges**: Very large ranges may cause timeouts (recommended: fewer than 10,000 cells).
3. **No Formulas**: Returns computed values only, not formula text.
4. **No Formatting**: Cell styles are not returned.
5. **Read-Only**: Does not modify the Excel file.

## Best Practices

### 1. Use the Most Specific Path

Include all subfolders in the path so the lookup resolves directly without ambiguity.

### 2. Verify the Site Name for Site Files

The site name must match the SharePoint **display name** (visible in the SharePoint admin centre), not the URL slug.
For example, if the URL is `https://tenant.sharepoint.com/sites/eng-team`, the display name may be `Engineering Team`.

### 3. Verify the Library Name for Site Files

The library name must match the **URL path segment** of the library, not necessarily its display name.
For example, the default library is `Shared Documents` in the URL but `Documents` in the display name — use `Shared Documents`.

### 4. Prefer Exact Ranges

Specify the tightest range that contains your data rather than entire columns to avoid unnecessary data transfer.

### 5. Handle Empty Cells

Check for empty string values in the returned rows and handle them appropriately in your workflow.

## Related Catalog Requests

- [Get Rows For Range](pages/GetRowsForRange.md) — Same operation using Drive ID + File ID (Shared Drive)
- [Get Rows For Range (Sheet)](pages/GetRowsForRangeSheet.md) — Same operation using File ID (Sheet area)
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md) — Obtain a Drive ID
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) — Obtain a File ID
- [Get Sheets Information](pages/GetSheetsInformation.md) — List worksheets in a file
- [Add Rows](pages/AddRows.md) — Write rows to a sheet

## Technical Implementation

### Classes

- **Catalog handler**: `SheetArea.getRowsForRangeByPath`
- **Path resolution**: `DriveService.getExcelFileByPath(sourcePath, lookupStrategy)`
- **Facade**: `SheetOperationsFacade.getGetRowsForRangeExtensionResponse(driveId, fileId, sheetName, range, allowRetry)`
- **Data retrieval**: `RowResponseService.getGetRowsForRangeExtensionResponse(driveId, fileId, sheetName, range, allowRetry)`

### Microsoft Graph API Endpoints Used

| Strategy   | Path Resolution Endpoint                                                                            | Data Endpoint                                                                  |
|------------|-----------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------|
| My Drive   | `GET /me/drive/root:/{path}`                                                                        | `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheet}/range(...)` |
| Site Files | `GET /sites?search={site}` → `GET /sites/{siteId}/drives` → `GET /drives/{driveId}/root:/{path}`  | `GET /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheet}/range(...)` |

### Permissions Required

- `Files.ReadWrite` — Access user's personal files
- `Sites.Read.All` — Required for Site Files strategy to search and access SharePoint sites

## Troubleshooting

### Issue: "File not found in Site Files at path: ..."

**Likely cause**: Site name or library name mismatch.

**Steps**:

1. Open the file in SharePoint and copy the URL.
2. The URL format is `https://tenant.sharepoint.com/sites/{siteSlug}/{libraryUrlSegment}/...`.
3. The **site display name** (not the URL slug) goes as the first path segment.
4. The **library URL segment** (not the display name) goes as the second path segment.

**Example**:

```
URL:         https://tenant.sharepoint.com/sites/eng-team/Shared%20Documents/Reports/Q1.xlsx
Site name:   Engineering Team    ← display name from SharePoint admin
Library:     Shared Documents    ← URL segment (decoded from Shared%20Documents)
Source Path: /Engineering Team/Shared Documents/Reports/Q1.xlsx
```

---

### Issue: "File not found in My Drive at path: ..."

**Likely cause**: The path doesn't match the actual folder structure in your OneDrive.

**Steps**:

1. Open OneDrive in a browser and navigate to the file.
2. The path shown in the breadcrumb (e.g. `My Files > Finance > Budget 2026.xlsx`) maps to `/Finance/Budget 2026.xlsx`.
3. Use that exact path as the Source Path.

---

### Issue: Empty Rows Returned

**Cause**: The range contains no data or the coordinates are outside the used area.

**Steps**:

1. Open the file and verify data is present in the specified range.
2. Check the range coordinates match the actual data position.

---

### Issue: Timeout on Large Range

**Cause**: The requested range is too large.

**Steps**:

1. Reduce the range to fewer than 10,000 cells.
2. Break the read into multiple smaller requests.

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) — Setup and configuration
- [Authentication](pages/Authentication.md) — Authentication requirements
- [Get Rows For Range](pages/GetRowsForRange.md) — Shared Drive variant (uses Drive ID + File ID)
- [Get Sheets Information](pages/GetSheetsInformation.md) — List available worksheets
