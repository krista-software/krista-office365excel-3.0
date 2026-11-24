# Release Notes

## Overview

This document provides version history, bug fixes, new features, and known issues for the Krista Office 365 Excel
Extension.

---

## Version 4.0.4

**Release Date**: June 2026

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 4.0.4        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.121      |
| Global Catalog Version     | GC-2026.07.1 |
| Microsoft Graph API        | v1.0         |

### Bug Fixes

#### **Fix: OAuth Consent Screen Not Shown for Quick Setup (AADSTS65001)**

**Problem**: When authenticating via the Quick Setup (Krista-managed Public OAuth) flow, the Microsoft consent
popup was closing automatically without showing the consent screen. Subsequent catalog requests failed with
Azure AD error `AADSTS65001` ("user has not consented to use the application").

**Root Causes Fixed**:

1. **Wrong scopes in auth URL** — The platform SDK's `buildAuthorizationUrl` for the Quick Setup flow was
   generating an auth URL containing platform-level mail scopes (`Mail.Send`, `Mail.ReadWrite`) instead of
   this extension's required `Files.ReadWrite` scope. When the user already had consent for those mail scopes,
   Microsoft silently issued the auth code without showing the consent screen, resulting in a token that
   lacked `Files.ReadWrite`. All subsequent Excel operations then failed with AADSTS65001.

   **Fix**: After `buildAuthorizationUrl`, the scope parameter is now checked for `Files.ReadWrite`. If
   missing, the scope is replaced with the correct Excel-specific scopes:
   `openid offline_access Files.ReadWrite User.Read`.

2. **prompt=consent not set** — The auth URL was using Microsoft's default `prompt=select_account`, which
   does not force the consent screen to appear when partial consent already exists.

   **Fix**: The `prompt` parameter is now always replaced with `prompt=consent` to ensure Microsoft always
   displays the full consent screen.

3. **MSAL4J error not handled** — `ExecutionException` from MSAL4J token refresh was caught generically,
   bypassing `AUTH_ERROR_RULES` entirely. Stale tokens were never removed and specific Azure AD error codes
   (e.g. AADSTS65001) were never surfaced to the user.

   **Fix**: `ExecutionException` is now caught separately, the real MSAL4J cause unwrapped, and
   `AUTH_ERROR_RULES` applied — removing the stale token and returning the correct user-facing error message.

4. **Callback returned HTML with `window.close()`** — The OAuth callback endpoint was returning an HTML page
   containing `<script>window.close();</script>`. Browsers silently block `window.close()` on tabs not opened
   via JavaScript `window.open()`, so the configuration panel was never populated after authentication.

   **Fix**: The callback endpoints now return plain text, matching the behaviour of other extensions
   (e.g. SharePoint). Auth completion is signalled server-side via `authorizationListener.authorized()`.

### Test Coverage

| Test Class                    | New Tests Added | Coverage                                                                      |
|-------------------------------|-----------------|-------------------------------------------------------------------------------|
| `AuthenticationResourceTest`  | 7               | All `testConnection` branches: null email, GraphServiceException, IOException, MAE delegated (scope override × 2, prompt replace/append × 2), finally block |

---

## Version 4.0.2

**Release Date**: June 2026

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 4.0.2        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.121      |
| Global Catalog Version     | GC-2026.06.3 |
| Microsoft Graph API        | v1.0         |

### What's New

#### **New Catalog Request — Get Rows For Range By Path**

A new **Sheet Operations** catalog request that retrieves rows from a specified range in an Excel worksheet by
resolving the file from its SharePoint or OneDrive **path** rather than requiring the user to supply a File ID.

**Request name**: `Get Rows For Range By Path`

**Inputs**:

| Parameter   | Type     | Notes                                                                               |
|-------------|----------|-------------------------------------------------------------------------------------|
| Source Path | Text     | Full path to the Excel file (e.g. `/Engineering/Shared Documents/Reports/Q1.xlsx`) |
| Sheet Name  | Text     | Worksheet name                                                                      |
| Range       | Text     | Excel range (e.g. `A1:D10`)                                                         |
| File Search | Pick One | `My Drive` or `Site Files`                                                          |
| Allow Retry | Boolean  | Optional retry on failure                                                           |

**Output**: `[ Composite ]` — identical to the existing *Get Rows For Range* request.

**File Search options**:

- **My Drive** — resolves the file directly from the authenticated user's personal OneDrive via
  `GET /me/drive/root:/{path}`. Use when you own the file.
- **Site Files** — resolves the file through the SharePoint Sites API (`GET /sites?search={siteName}` →
  drive lookup → `GET /drives/{driveId}/root:/{itemPath}`). Use for SharePoint site libraries and
  Microsoft Teams channel files. The first path segment must be the site display name, the second must
  be the document library URL segment (e.g. `Shared Documents`).

**Path behaviour**:

- Leading slash is optional and normalised automatically.
- URL-encoded characters (e.g. `Shared%20Documents`) are decoded before lookup.
- Folder and filename matching is case-insensitive.
- An unrecognised File Search value throws `IllegalArgumentException` with a clear message listing valid options.

**Documentation**: [Get Rows For Range By Path](pages/GetRowsForRangeByPath.md)

### Test Coverage

A new test class `DriveServiceGetExcelFileByPathTest` covers both File Search options and all error paths:

| Test Group          | Scenarios covered                                                                                       |
|---------------------|---------------------------------------------------------------------------------------------------------|
| Input Validation    | null path, blank path, missing filename, null File Search, blank File Search, Site Files path too short |
| Unknown File Search | unrecognised value → `IllegalArgumentException` with "Unknown File Search" message                      |
| Path Parsing        | leading slash stripped, URL-decoded path (`%20`), nested subfolder                                      |
| My Drive            | file found, item not found (null), `GraphServiceException`, `IOException`                               |
| Site Files          | path too short, `IOException`                                                                           |

**Total**: 18 tests, all passing.

---

## Version 4.0.1

**Release Date**: June 2026

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 4.0.1        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.121      |
| Global Catalog Version     | GC-2026.06.2 |
| Microsoft Graph API        | v1.0         |

### What's New

#### **New Catalog Request: Search All Shared Excel Files in Drive**

A new Shared Drive catalog request that recursively searches for files of a selected format across **all folders and sub-folders** within a specified drive using the Microsoft Graph search API.

- **Area**: Shared — **Type**: QUERY_SYSTEM
- **Inputs**:
  - **Drive Id** (Text, required) — the drive to search in; obtain from *Get All Shared Drives Information*
  - **File Format** (PickOne, required) — select one of: `.xlsx`, `.xls`, `.csv`, `.tsv`
- **Output**: `Files` — list of Excel File entities (Drive ID, File ID, File Name, Web URL)
- Performs server-side recursive traversal via `GET /me/drives/{driveId}/root/search(q='{extension}')` — no folder path is needed
- Automatically follows `@odata.nextLink` pagination cursors to retrieve all result pages
- Duplicate items returned across pages are deduplicated automatically

> **When to use this vs. Get Shared Excel Files**: Use *Search All Shared Excel Files in Drive* when files may be stored in sub-folders inside a known drive. Use *Get Shared Excel Files* when you only need root-level files or need to discover Drive IDs across all accessible shared drives.

**Documentation**: [Search Excel Files in Drive](pages/SearchExcelFilesInDrive.md)

### Documentation Updates

#### **Get Shared Excel Files — Root-Level Behaviour Clarified**

Updated documentation to explicitly state that when a Drive ID is provided, *Get Shared Excel Files* returns only files located in the **root folder** of that drive — files inside sub-folders are not included. This is existing behaviour; the documentation now clearly documents this constraint and directs users to *Search All Shared Excel Files in Drive* for recursive searches.

---

## Version 4.0.0

**Release Date**: June 2026

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 4.0.0        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.121      |
| Global Catalog Version     | GC-2026.06.2 |
| Microsoft Graph API        | v1.0         |

### What's New

#### **Authentication System Rewrite — Microsoft Auth SDK**

This release replaces the ScribeJava-based OAuth implementation with the Krista Microsoft Auth SDK. The auth tab now exposes three clearly separated authentication modes and natively supports MCP (multi-user, per-user token isolation).

#### **Three Authentication Modes**

- **Quick Setup** (`AUTH_CODE_PUBLIC`): Uses Krista's pre-registered Azure AD app. Only an email address is required. Ideal for individual users and small teams with no Azure AD configuration.
- **User Authentication** (`AUTH_CODE_PRIVATE`): OAuth 2.0 authorization code flow using the organisation's own Azure AD app registration. Requires Client ID, Client Secret, and Tenant ID. Grants delegated access scoped to the signed-in user.
- **Server to Server** (`CLIENT_CREDENTIALS_SECRET`): OAuth 2.0 client credentials flow — no user sign-in. Designed for automated workflows and background processes that require unattended access.

#### **MCP (Multi-User) Per-User Token Isolation**

- Each Krista user now gets a hermetic credential and token keyspace derived from their invoker ID.
- Tokens are stored and refreshed independently per user — one user's token expiry or re-auth does not affect other users.
- Auth callbacks route to the originating user's context via `invokeAsUser`-first resolution with appliance fallback.
- Services (`GraphServiceClient`, token store) are initialised lazily on first use per user, eliminating startup-time failures for unauthenticated users.

#### **Audit Dashboard**

- All authentication and catalog operations are now recorded in a persistent audit log backed by the KV store.
- Each record shows event type, status, reason, triggered-by user, duration, and execution logs.
- Records are auto-archived at 2,000 entries; the archive is browsable from the audit endpoint.
- Auth error counters (token expiry, invalid secret, config failure) are reset automatically on successful credential save.

---

### Bug Fixes

#### **Fixed: Legacy Private Tab Credentials Migrating to Server to Server**

- **Root Cause**: When loading stored credentials that were saved with the old `authType=Private` tab (pre-SDK migration), the load and create paths defaulted the missing `authFlow` to `CLIENT_CREDENTIALS_SECRET` (Server to Server). The old Private tab used OAuth authorization code flow (delegated, user signs in with their own Azure AD app) — which is equivalent to the new User Authentication mode, not Server to Server.
- **Fix**: `OfficeExcelAttributeStore.load()` and `OfficeExcelAttributes.create()` now detect the legacy case (`authType=Private` + no `authFlow` stored) and migrate it to `authType=Public` + `AUTH_CODE_PRIVATE` (User Authentication). Connections that explicitly store `authFlow=CLIENT_CREDENTIALS_SECRET` continue to load as Server to Server.
- **Impact**: Users whose connections were saved under the old Private tab will be correctly presented as User Authentication after upgrade, retaining their Client ID, Client Secret, Tenant ID, and email credentials.

#### **Fixed: "No Execution Logs Captured" in Audit Records**

- **Root Cause**: `Office365ExcelAuditService.startRequest(invokerId, catalogName)` called `startRequest()` (which opens the Log4j2 capture buffer) and then immediately called `record(TRIGGERED)`. The SDK's `record()` method calls `RequestLogCapture.stopCapture()` for any `_TRIGGERED` event, which destroyed the capture buffer before any real work logs were emitted. All subsequent `LOGGER` calls in the request had no active capture, so the final SUCCESS/FAILURE record always showed "No execution logs captured for this request."
- **Fix**: Swapped the call order — `record(TRIGGERED)` is now recorded first (before capture starts, so it correctly holds no logs), then `startRequest()` opens a fresh capture window for the duration of the actual work.
- **Impact**: Execution logs now appear in all SUCCESS and FAILURE audit records for OAuth Callback, Test Connection, Save Credentials, and all catalog operations.

---

### Migration Notes

#### **For Users Upgrading from 3.0.x**

- **Old Public tab** → loads as **Quick Setup** (no change in behaviour).
- **Old Private tab** (own Azure AD app, auth code flow) → automatically migrates to **User Authentication**. Credentials are preserved; no re-entry required.
- **New Server to Server connections** must be configured explicitly in the new auth tab. No existing connections are silently reclassified as S2S.

---

## Version 3.0.3

**Release Date**: April 2026

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 3.0.3        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.121      |
| Global Catalog Version     | GC-2026.04.4 |
| Microsoft Graph API        | v1.0         |

### What's New

## [KE-3154] - New Async Get Rows For Range, Allow Retry Expansion & Reliability Fixes

This release introduces a new asynchronous catalog request for Get Rows For Range, expands the Allow Retry
(error handling) feature to 13 additional SharedArea catalog requests, and fixes critical reliability issues
in the Graph API client and validation pipeline.

### New Features

#### **New Catalog Request: Get Rows for Range Async**

- Added `getRowsForRangeAsync` — an asynchronous version of the existing Get Rows For Range request
- Executes the Graph API range read on a background thread (`ExecutorService`) and returns a Task ID immediately
- Added companion `getRowsForRangeResult` (Wait-for-Event) request to retrieve the async result using the Task ID
- Enables non-blocking reads for large Excel ranges that previously caused timeouts in synchronous mode
- Thread-safe: propagates `ThreadLocalProxy` context to the background thread and cleans up on completion

#### **Allow Retry (Error Handling) Expanded to SharedArea Requests**

- Added `@Field.Boolean(name = "Allow Retry")` parameter to 13 SharedArea catalog requests that previously
  had no error handling support:
  - Get Rows For Range, Search, Get Cell Data, Get Shared Excel File, Convert To PDF,
    Delete Column, Add Row With Index, Update Row With Index, Delete Row With Index,
    Get Complete Row Info, Add Multiple Cell Data, Get Column Values For Keys In Sheet,
    Clear Shared Sheet For Range
- When `allowRetry=true`, these requests now validate inputs (drive, file, sheet) before executing and
  present interactive retry prompts on validation failure

### Bug Fixes

#### **Fixed: OkHttp Timeout Causing 403 "Could not obtain a WAC access token"**

- **Root Cause**: The default OkHttp read timeout (10s) was too short for `createSession` calls on large
  workbooks (which take 25-30s), causing the session manager to return null. The subsequent sessionless
  Graph API call then failed with a 403 "Could not obtain a WAC access token" error.
- **Fix**: Configured custom OkHttp timeouts on the `GraphServiceClient` — **60s read timeout** and
  **120s call timeout** — giving large workbook sessions sufficient time to initialize.
- **Impact**: Eliminates the 403 WAC token error on large workbooks caused by premature timeout.

#### **Fixed: 403 Access Denied Errors Misreported as Validation Failures**

- **Root Cause**: When `ValidationOrchestrator.validate()` encountered an `IllegalStateException`
  (thrown on 403 access denied or WAC token errors), it was caught by the generic `Exception` handler
  and presented as a retryable validation error — misleading the user.
- **Fix**: Added explicit `IllegalStateException` catch block in `ValidationOrchestrator.validate()` that
  propagates the exception to the caller instead of swallowing it.
- **Fix**: Added `IllegalStateException` handling in `AreaOperationHandler.execute()` that catches the
  propagated exception and returns a proper access error response with remediation action.
- **Impact**: Users now see a clear "access error" message instead of a misleading "validation failed" prompt.
---

## Version 3.0.2

**Release Date**: April 2026

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 3.0.2        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.121      |
| Global Catalog Version     | GC-2026.04.3 |
| Microsoft Graph API        | v1.0         |

### What's New
## [KE-3154] - WAC Access Token & Session Retry Fix
This patch release fixes the `403 AccessDenied – Could not obtain a WAC access token` error that occurred when
accessing large Excel workbooks on shared drives. It also improves session creation resilience and corrects
misleading error messages.

### Bug Fixes

#### **Resolved: 403 "Could not obtain a WAC access token" Error**

- **Root Cause**: When the cached workbook session's WAC token expired, the previous retry logic discarded the
  session and retried without one. Sessionless calls require Microsoft Graph to acquire a WAC token on-the-fly,
  which frequently fails on large workbooks — producing the same 403 error on every retry.
- **Fix**: The extension now detects the WAC token error (`403 + "WAC access token"` in the message) and creates a
  **fresh session** (via `POST /workbook/createSession`) instead of retrying without a session. The new session
  obtains a fresh WAC token during creation, and subsequent API calls reuse it reliably.
- **Impact**: Eliminates the persistent 403 loop that made large shared-drive workbooks inaccessible.

#### **Resolved: Misleading "Sheet not found" Error on 403/500 Responses**

- **Root Cause**: Non-404 errors (403 access denied, 500 server errors) were incorrectly reported as
  "Sheet does not exist" in validation and range-read flows.
- **Fix**: Error handling now properly distinguishes 404 (sheet not found), 403 (access denied), and other HTTP
  errors, surfacing the actual error code and message to the user.

### Reliability Improvements

#### **Session Creation Retry with Exponential Backoff**

- Session creation (`POST /workbook/createSession`) now retries up to 3 times with exponential backoff
  (1s → 2s → 4s) on transient errors (HTTP 429, 5xx).
- Non-retryable errors (4xx other than 429) fail immediately without unnecessary retries.

#### **Fresh Session Retry on Session Rejection**

- Added `refreshSessionId()` and `refreshSessionHeaderOptions()` methods to `WorkbookSessionManager` that
  invalidate the cached session and create a new one in a single call.
- All session-rejection retry paths (`worksheetExists`, `getDataForRange`) now use the fresh-session approach
  instead of falling back to sessionless calls.

#### **Proper 403 Propagation in Sheet Name Validation**

- `SheetNameValidator` now catches and re-throws `IllegalStateException` (thrown on 403 access denied) instead
  of silently returning `false` and misleading the user with "sheet doesn't exist".

### Affected Catalog Requests

All catalog requests that validate sheet names or read sheet data are covered, including: Get Rows For Range,
Get Cell Value, Get Cell Data, Add Data To Cell, Add Multiple Cell Data, Add Formula To Cell, Add Rows,
Add/Update/Delete Row With Index, Get Complete Row Info, Get Column Values For Keys, Delete Column,
Clear Sheet Range, Delete Sheet, Search, VLookup, Create Table For Range — and all "Shared File" variants.

---

## Version 3.0.1

**Release Date**: April 2026

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 3.0.1        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.121      |
| Global Catalog Version     | GC-2026.04.3 |
| Microsoft Graph API        | v1.0         |

### What's New
## [KE-3154] -
This patch release focuses on reliability and performance for reading data from large Excel workbooks. It resolves
intermittent timeouts, empty-range failures, and misleading "Sheet not found" errors reported on files with millions
of cells.

### Reliability & Performance Improvements

#### **Faster, More Reliable Reads on Large Workbooks**

- **Workbook Session Caching**: The extension now opens a Microsoft Graph workbook session on the first request to a
  file and reuses it across subsequent operations on the same workbook, with an in-memory cache that refreshes every
  5 minutes.
- **Benefit**: Eliminates repeated cold workbook loads that previously caused `MaxRequestDurationExceeded` timeouts
  (often surfaced as "Sheet not found") on large Excel files. After the first request to a workbook warms the session,
  follow-up reads and writes on the same file complete in under a second.
- **Automatic recovery**: If the server evicts a session, the extension transparently retries the request without a
  session so the user never sees an "invalid session" error.


#### **Improved Troubleshooting Diagnostics**

- Added structured, single-line log entries for range reads and session activity, including the outgoing Graph
  request URL, session cache hit/miss, response shape (row/column counts, presence of text and values), and the
  Graph error code / HTTP status on failures.
- **Benefit**: Dramatically faster root-cause analysis for customer-reported issues — "sheet not found", timeout,
  and permission errors can now be distinguished from the first log line without enabling debug-level logging.

#### **Affected Catalog Requests**

The reliability improvements above apply to every catalog request that reads sheet data or validates a sheet name,
including: Get Rows For Range, Get Cell Value, Get Cell Data, Add Data To Cell, Add Multiple Cell Data,
Add Formula To Cell, Add Rows, Add/Update/Delete Row With Index, Get Complete Row Info, Get Column Values For Keys,
Delete Column, Clear Sheet Range, Delete Sheet, Search, VLookup, Create Table For Range — and all "Shared File"
variants of these operations.

---

## Version 3.0.0

**Release Date**: November 2025

### Version Information

| Component                  | Version      |
|----------------------------|--------------|
| Extension Version          | 3.0.0        |
| Developer                  | Simran Sethi |
| Krista Service APIs (Java) | 1.0.115      |
| Global Catalog Version     | GC-2025.11.1 |
| Microsoft Graph API        | v1.0         |

### What's New

#### **🔐 New Authentication System - Dual Authentication Modes**

This major release introduces a completely redesigned authentication experience with two distinct authentication modes to suit different use cases.

#### **New Authentication Tabs**

- **Public Authentication Tab**:
  - Simple, user-friendly authentication for individual users
  - No Azure AD configuration required
  - Pre-configured Microsoft application credentials
  - Perfect for personal OneDrive and SharePoint access
  - User-delegated permissions - no admin consent needed
  - Quick setup - just enter your email and authenticate
  - Ideal for individual users and small teams

- **Private Authentication Tab**:
  - Enterprise-grade authentication for organizations
  - Custom Azure AD application support
  - Full control over permissions and scopes
  - Tenant-specific authentication
  - Organization-wide file and site access
  - Perfect for enterprise deployments and service accounts
  - Supports advanced scenarios requiring elevated permissions

#### **Enhanced User Experience**

- **Intuitive UI**: Clean, modern interface with separate tabs for each authentication mode
- **Smart Defaults**: Public auth comes pre-configured for immediate use
- **Flexible Configuration**: Private auth allows full customization for enterprise needs
- **Seamless Switching**: Easy to switch between authentication modes based on your requirements

#### **Public Authentication Improvements**

- **No Admin Approval Required**:
  - User-delegated permissions eliminate the need for admin consent
  - No more "Need admin approval" or "unverified app" warnings
  - Instant authentication for standard operations

- **Optimized Permissions**:
  - `Files.ReadWrite` - Access to user's files and shared files
  - `Sites.Read` - Read access to user's SharePoint sites
  - `User.ReadWrite` - User profile access
  - `openid` and `offline_access` - OAuth essentials

- **Streamlined Flow**:
  - Simplified permission model for better user experience
  - Faster authentication process
  - Reduced friction for new users
  - All core features work seamlessly

#### **Private Authentication Features**

- **Custom Azure AD Apps**: Use your organization's registered applications
- **Tenant-Specific**: Authenticate against specific Azure AD tenants
- **Advanced Permissions**: Support for organization-wide scopes like `Files.ReadWrite.All` and `Sites.Read.All`
- **Service Account Support**: Perfect for automated workflows and background processes
- **Full Control**: Complete control over client credentials and permissions

### Migration Notes

#### **For Existing Users**

- **Public Auth Users**: No action required - authentication will work seamlessly without admin approval
- **Private Auth Users**: No changes - private authentication continues to work as before
- **Permissions**: User-delegated permissions provide access to all files and sites the user can access

#### **For Administrators**

- **Admin Consent**: No longer required for public authentication
- **Security**: User-delegated permissions are more secure as they're scoped to individual users
- **Deployment**: Easier rollout to end users without IT intervention

### Known Limitations

- **Organization-Wide Access**: If you need to access files across the entire organization (not just user's files), use private authentication with appropriate service account permissions
- **Scope Changes**: Applications requiring `.All` permissions should continue using private authentication mode

---

## Version 2.0.28

**Release Date**: October 2025

### Version Information

| Component                  | Version           |
|----------------------------|-------------------|
| Extension Version          | 2.0.28            |
| Developer                  | Vaibhav Choudhary |
| Krista Service APIs (Java) | 1.0.115           |
| Global Catalog Version     | GC-2025.10.3      |
| Microsoft Graph API        | v1.0              |

### What's New

#### **Major Error Message Enhancement Initiative**

This release focuses entirely on improving user experience through comprehensive error message enhancements across the
entire extension.

#### **User-Friendly Error Messages**

- **Validator Messages**: Completely rewrote all validation error messages to be more helpful and actionable
- **Authentication Errors**: Enhanced authentication error messages with clear next steps and administrator guidance
- **Service Layer Errors**: Improved error messages in WorkbookService and WorksheetService with better context
- **Consistent Tone**: All error messages now use friendly, helpful language instead of technical jargon

#### **Enhanced Documentation**

- **New Troubleshooting Guide**: Comprehensive troubleshooting documentation with solutions for common issues
- **Error Handling Section**: Added detailed error handling guidance to main README
- **Common Error Codes**: Reference table for quick error resolution
- **Self-Service Support**: Users can now resolve most issues independently

### Detailed Improvements

#### **Validator Error Messages Enhanced**

**File and Resource Validators**:

- **FileIdValidator**: Now explains how to find file IDs and provides permission guidance
- **DriveIdValidator**: Clarifies how to obtain drive IDs using "Get All Shared Drives Information"
- **FileNameValidator**: Lists all supported Excel extensions with descriptions

**Sheet and Table Validators**:

- **SheetNameValidator**: Explains Excel naming rules and suggests using "Get Sheets Information"
- **TableNameValidator**: Emphasizes case-sensitivity and provides clear accessibility guidance
- **ColumnNameValidator**: Highlights case-sensitivity with helpful examples

**Data Format Validators**:

- **CellIdValidator**: Shows proper Excel cell reference formats (A1, B2, AA10)
- **RangeValidator**: Provides multiple valid Excel range format examples
- **RowIndexValidator**: Clarifies Excel row numbering starting from 1
- **FormulaValidator**: Includes formula examples and mentions 8,192 character limit

#### **Authentication Error Messages Improved**

**Enhanced MicrosoftAuthErrorConstants**:

- **Session Expired**: Clear re-authorization guidance
- **Permission Issues**: Specific administrator contact instructions
- **Account Problems**: Detailed explanations for disabled/deleted accounts
- **Application Issues**: Clear guidance for Azure AD configuration problems
- **Service Unavailable**: Helpful retry and escalation guidance

#### **Service Layer Enhancements**

**WorkbookService Improvements**:

- Better workbook not found messages with permission context
- Enhanced batch processing error messages with data format guidance
- Improved CSV processing errors with file format validation tips
- More helpful formula execution errors with syntax guidance

**WorksheetService Improvements**:

- Clearer worksheet deletion error messages
- Enhanced sheet not found errors with troubleshooting steps
- Better permission-related error explanations

#### **Documentation Enhancements**

**New Troubleshooting Guide** (`pages/Troubleshooting.md`):

- **Authentication Issues**: Step-by-step resolution for all auth problems
- **File Access Issues**: Comprehensive file permission and access guidance
- **Data Format Issues**: Clear examples of correct formats for all data types
- **Formula Issues**: Excel formula troubleshooting with examples
- **Performance Issues**: Guidance for handling large datasets and timeouts
- **Error Code Reference**: Quick lookup table for common error codes

**Updated README.md**:

- **Error Handling Section**: New comprehensive error handling guidance
- **Common Error Types**: Categorized error types with solutions
- **Getting Help Section**: Enhanced with troubleshooting guide reference

### User Experience Improvements

#### **Key Benefits**

1. **Faster Problem Resolution**: Users can quickly understand and fix issues
2. **Reduced Support Burden**: Clear guidance reduces administrator intervention needs
3. **Better Learning Experience**: Error messages help users understand Excel integration
4. **Self-Service Capability**: Comprehensive troubleshooting enables independent problem-solving
5. **Consistent Experience**: All error messages follow the same helpful, friendly tone

#### **Error Message Principles Applied**

- **Clear and Concise**: No technical jargon, plain English explanations
- **Actionable Guidance**: Every error includes specific steps to resolve
- **Helpful Context**: Users understand what went wrong and why
- **Examples Provided**: Correct formats and usage examples included
- **Appropriate Tone**: Helpful and supportive, never blaming

### Breaking Changes

None in this version. All existing functionality remains unchanged.

### Deprecations

None in this version.

### Security Updates

No security changes in this version.

### Performance Improvements

No performance changes in this version.

### API Changes

No API changes in this version. All existing catalog requests remain fully compatible.

---

## Version 2.0.26

**Release Date**: October 2025

### Version Information

| Component                  | Version          |
|----------------------------|------------------|
| Extension Version          | 2.0.26           |
| Developer                  | Vrushali Gaikwad |
| Krista Service APIs (Java) | 1.0.115          |
| Global Catalog Version     | GC-2025.10.2     |
| Microsoft Graph API        | v1.0             |

### What's New

#### Documentation Overhaul

- Complete documentation rewrite following comprehensive guidelines
- Individual pages for each catalog request with detailed information
- Enhanced authentication documentation with OAuth 2.0 flows and token management
- Comprehensive configuration guide with troubleshooting
- Organized sidebar navigation by functional categories

#### Enhanced Features

- Improved retry mechanism for handling rate limiting
- Better error messages and validation
- Support for larger datasets in batch operations

### Resolved Bugs

#### KE-2394: Entity Layout Not Honored

- **Issue**: Office 365 Excel does not honor entity layout when writing to spreadsheet
- **Impact**: Data written to Excel may not match expected column order
- **Resolution**: Fixed entity layout processing to maintain correct column order
- **Status**: Resolved

#### KE-2500: Retry and Rate Limiting

- **Issue**: Extension does not properly handle rate limiting from Microsoft Graph API
- **Impact**: Operations fail when rate limits are exceeded
- **Resolution**: Implemented automatic retry with exponential backoff for rate limiting (429 errors)
- **Status**: Resolved

#### KE-2504: Unable to Add Larger Datasets

- **Issue**: Extension fails when attempting to add large datasets to Excel
- **Impact**: Bulk data operations fail or timeout
- **Resolution**: Optimized batch processing and increased timeout limits
- **Status**: Resolved

### Known Issues

#### KE-1226: DateTime Display Issue

- **Issue**: Time is not displayed inside the DateTime range field in "Add Rows With List Of Entity" request
- **Impact**: Time component of DateTime values may not be visible in Excel
- **Workaround**: Format DateTime values as text before adding to Excel
- **Status**: Under Investigation
- **Planned Fix**: Version 2.1.0

### Breaking Changes

None in this version.

### Deprecations

None in this version.

### Security Updates

- Updated OAuth 2.0 token handling for improved security
- Enhanced credential storage encryption
- Improved error messages to avoid exposing sensitive information

### Performance Improvements

- Optimized batch operations for large datasets
- Reduced API call overhead through better caching
- Improved retry logic to minimize delays

### API Changes

No API changes in this version. All existing catalog requests remain compatible.

---

## Version History

### Version 2.0.23

**Release Date**: September 2025

**Changes**:

- Initial support for asynchronous operations
- Added batch processing capabilities
- Bug fixes and stability improvements

### Version 2.0.22

**Release Date**: November 2024

**Changes**:

- Enhanced error handling
- Improved authentication flow
- Performance optimizations

### Version 2.0.21

**Release Date**: October 2024

**Changes**:

- Added support for additional Excel operations
- Bug fixes
- Documentation updates

---

## Upgrade Instructions

### Upgrading from 2.0.23 to 2.0.24

1. **Backup Configuration**:
    - Export current extension configuration
    - Document any custom settings

2. **Update Extension**:
    - Navigate to Extensions in Krista
    - Select Office 365 Excel Extension
    - Click "Update" to version 2.0.24

3. **Verify Configuration**:
    - Review extension configuration
    - Test connection
    - Validate authentication

4. **Test Workflows**:
    - Test existing workflows using the extension
    - Verify all catalog requests function correctly
    - Check error handling and retry behavior

5. **Review Documentation**:
    - Review updated documentation for new features
    - Update any custom documentation or training materials

### No Action Required

- Existing configurations remain compatible
- No changes to catalog request signatures
- Authentication settings are preserved

---

## Compatibility

### Supported Platforms

| Platform            | Version | Status    |
|---------------------|---------|-----------|
| Krista Appliance    | 3.5.1+  | Supported |
| Microsoft Graph API | v1.0    | Supported |
| Office 365 Excel    | Online  | Supported |
| Microsoft Entra ID  | v2.0    | Supported |

### Browser Compatibility (for OAuth)

| Browser | Version | Status    |
|---------|---------|-----------|
| Chrome  | Latest  | Supported |
| Firefox | Latest  | Supported |
| Edge    | Latest  | Supported |
| Safari  | Latest  | Supported |

### Required Permissions

- **Microsoft Graph API**:
    - `Files.ReadWrite` - Read and write user's files
    - `Files.ReadWrite.All` - Read and write all files user can access
    - `openid` - OpenID Connect authentication
    - `offline_access` - Refresh token support

---

## Support

### Getting Help

If you encounter issues with this version:

1. **Review Documentation**:
    - [Extension Configuration](pages/ExtensionConfiguration.md)
    - [Authentication](pages/Authentication.md)
    - [Troubleshooting sections in catalog request pages](../README.md#catalog-requests)

2. **Check Known Issues**:
    - Review known issues section above
    - Check if your issue is already documented

3. **Contact Support**:
    - Provide extension version (2.0.27)
    - Include error messages and logs
    - Describe steps to reproduce the issue

### Reporting Bugs

When reporting bugs, please include:

- Extension version: 2.0.24
- Krista Appliance version
- Catalog request being used
- Complete error message
- Steps to reproduce
- Expected vs actual behavior

---

## Feedback

We value your feedback! Please share:

- Feature requests
- Documentation improvements
- Usability suggestions
- Bug reports

Contact your Krista support representative or submit feedback through the Krista platform.

---

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md) - Setup and configuration guide
- [Authentication](pages/Authentication.md) - OAuth 2.0 authentication and token management
- [Creating Microsoft App](pages/CreatingMicrosoftApp.md) - Microsoft Entra ID application setup
- [Catalog Requests](../README.md#catalog-requests) - Complete list of available operations