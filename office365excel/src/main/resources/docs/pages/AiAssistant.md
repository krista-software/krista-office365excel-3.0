# AI Assistant (MCP)

## Overview

The AI Assistant tab enables natural language interaction with Excel workbooks stored in OneDrive. Instead of configuring catalog requests manually, you can ask questions and give instructions in plain English — the extension translates them into Microsoft Graph API calls automatically.

Both the **NQL tab** in Krista Studio and **Krista Agent / MCP client** tool calls use the same pipeline: the AI determines the exact Graph API call needed (`method`, `path`, `body`), and the extension executes it generically. No hardcoded intents — any valid Graph workbook endpoint works without code changes.

---

## Setup

### Step 1 — Configure Excel Authentication

The AI Assistant uses the same OneDrive/Excel connection as all other catalog requests. Complete the **Authentication** tab first (see [Authentication](pages/Authentication.md)).

### Step 2 — Configure an AI API Key

1. Open the extension's **AI Assistant** tab in Krista Studio.
2. Go to **AI Settings**.
3. Select a provider and paste in the API key:

| Provider | Model Used | Where to Get the Key |
|---|---|---|
| **Anthropic (Claude)** | claude-haiku-3-5 | [console.anthropic.com](https://console.anthropic.com) → API keys |
| **OpenAI (GPT)** | gpt-4o-mini | [platform.openai.com](https://platform.openai.com) → API keys |
| **Google (Gemini)** | gemini-2.5-flash | [aistudio.google.com](https://aistudio.google.com) → Get API key |

4. Click **Save**.

If no AI key is configured, queries return:
> _AI API key is not configured. Please set it in the AI Assistant tab → AI Settings._

---

## How It Works

Every natural language question goes through the same pipeline:

```
User question
  → AI plans the Graph API call  { method, path, body }
  → {itemId} placeholder resolved automatically from filename in question
  → Extension executes the call against Microsoft Graph
  → Response converted to table rows (noisy metadata fields filtered out)
  → Columns derived dynamically from response fields
  → GENERATED KQL field shows the executed method + path
```

The AI returns one of four HTTP methods:

| Method | When Used |
|---|---|
| `GET` | Read files, list files, read cell values, list worksheets, list tables |
| `POST` | Search files, add a worksheet, clear a cell range |
| `PATCH` | Write cell values, update a range |
| `DELETE` | Delete a worksheet |

---

## Supported Operations

### Finding and Listing Files

| Operation | Example Question | Graph Call |
|---|---|---|
| All Excel files (own + shared) | "Show all my Excel files" | `POST /search/query` with `FileType:xlsx OR FileType:xls OR FileType:xlsm` |
| Files shared with me | "Show files shared with me" | `GET /me/drive/sharedWithMe?$top=50` |
| Recent files | "Show my recent Excel files" | `GET /me/drive/recent?$top=50` |
| Files in root | "List Excel files in my drive" | `GET /me/drive/root/children?$top=50` |
| Search by name | "Find budget spreadsheets" | `POST /search/query` with KQL `budget FileType:xlsx` |

### Reading Workbook Data

All read operations require a filename in your question. The extension resolves the file automatically — you do not need to know the file ID.

| Operation | Example Question | Graph Call |
|---|---|---|
| List worksheets | "List the sheets in Budget.xlsx" | `GET /me/drive/items/{itemId}/workbook/worksheets` |
| Read a cell | "Get the value of A1 in Sheet1 in Budget.xlsx" | `GET /me/drive/items/{itemId}/workbook/worksheets/Sheet1/range(address='A1')` |
| Read a range | "Read cells A1 to C5 in Sheet1 in Budget.xlsx" | `GET /me/drive/items/{itemId}/workbook/worksheets/Sheet1/range(address='A1:C5')` |
| Read all used data | "Read all data in Sheet1 in Budget.xlsx" | `GET /me/drive/items/{itemId}/workbook/worksheets/Sheet1/usedRange` |
| List tables | "List tables in Budget.xlsx" | `GET /me/drive/items/{itemId}/workbook/tables` |
| Get table rows | "Get rows from Table1 in Budget.xlsx" | `GET /me/drive/items/{itemId}/workbook/tables/Table1/rows` |

### Writing Cell Values

| Operation | Example Question | Graph Call |
|---|---|---|
| Write a single cell | "Add 'Hello' to cell A1 in Sheet1 in Budget.xlsx" | `PATCH /me/drive/items/{itemId}/workbook/worksheets/Sheet1/range(address='A1')` |
| Write a row of values | "Set A1 to 'Name', B1 to 'Score' in Sheet1 in Budget.xlsx" | `PATCH .../range(address='A1:B1')` |
| Write multiple rows | "Write data to A1:B2 in Sheet1 in Budget.xlsx" | `PATCH .../range(address='A1:B2')` |

The body the AI generates for PATCH writes:

```json
{ "values": [["CellValue"]] }
```

For a range of cells, values is a 2D array where each inner array is a row:

```json
{ "values": [["Name", "Score"], ["Alice", "95"]] }
```

A successful write returns:
> _Cell updated successfully. (Sheet1!A1)_

### Worksheet Management

| Operation | Example Question | Graph Call |
|---|---|---|
| Add a worksheet | "Add a sheet called Summary to Budget.xlsx" | `POST /me/drive/items/{itemId}/workbook/worksheets/add` |
| Delete a worksheet | "Delete the sheet named OldData from Budget.xlsx" | `DELETE /me/drive/items/{itemId}/workbook/worksheets/OldData` |
| Clear a cell range | "Clear cell A1 in Sheet1 in Budget.xlsx" | `POST .../range(address='A1')/clear` |

A successful deletion or no-body action returns:
> _Action completed successfully._

---

## Automatic File ID Resolution

All workbook operations require an internal OneDrive item ID in the URL. The AI generates a path like `/me/drive/items/{itemId}/workbook/...` and the extension automatically resolves `{itemId}` before executing — you only need to name the file in your question.

**Resolution process:**

1. The extension extracts the filename from your question using a pattern match for `.xlsx`, `.xls`, `.xlsm`, `.xlsb` extensions.
2. It calls `GET /me/drive/root/search(q='filename')&$top=10&$select=id,name`.
3. It prefers an exact name match; falls back to the first result if no exact match.
4. The real item ID is substituted into the path.
5. The call is executed with the resolved URL.

**If the file is not found**, the query returns:
> _Could not find the Excel file mentioned in your question. Please check the file name and try again._

**Tips for reliable resolution:**
- Include the full filename with extension: `Budget.xlsx` not just `Budget`
- Filenames are matched case-insensitively
- If multiple files share the same name, the first search result is used — be specific if needed

---

## Result Columns

Columns are **fully dynamic** — derived at runtime from whatever fields the Graph API returns. There is no fixed column list. Noisy internal metadata fields are automatically filtered out.

### File Search Results

| Column | Graph Field | Description |
|---|---|---|
| Name | `name` | Filename |
| Last Modified | `lastModifiedDateTime` | Date last modified (date only) |
| Modified By | `lastModifiedBy.user.displayName` | Display name of the last modifier |
| Created | `createdDateTime` | Date created |
| Size | `size` | File size (formatted: B / KB / MB) |
| Folder | `parentReference.name` | Parent folder name |
| Link | `webUrl` | URL to open the file in OneDrive |
| File Id | `id` / `hitId` | Internal OneDrive item ID |

### Cell / Range Read Results

| Column | Source | Description |
|---|---|---|
| Cell | `address` | Cell or range address (e.g. `Sheet1!A1`) |
| Value | `values` | Flattened cell contents — columns separated by ` \| `, rows by ` / ` |

For example, reading `A1:B2` with values `[["Name","Score"],["Alice","95"]]` displays:
```
Value: Name | Score / Alice | 95
```

The following range metadata fields are automatically hidden from results as they add no useful information: `columnHidden`, `rowHidden`, `columnIndex`, `rowIndex`, `cellCount`, `columnCount`, `rowCount`, `hidden`, `numberFormat`, `formulas`, `text`, and similar internal fields.

---

## KQL Reference (File Search)

When the question is about finding files, the AI generates a KQL query for `POST /search/query`. Common patterns:

| KQL | Matches |
|---|---|
| `FileType:xlsx` | Excel workbook files |
| `Title:"Sales Report"` | Files with that exact name phrase |
| `Author:"John Smith"` | Files created by that person |
| `LastModifiedTime>2024-01-01` | Files modified after a date |
| `LastModifiedTime>today-7` | Files modified in the last 7 days |
| `Path:"finance"` | Files in a path containing "finance" |
| `budget AND FileType:xlsx` | Excel files with "budget" in the name or content |

The search is automatically scoped to Excel file types (`xlsx`, `xls`, `xlsm`, `xlsb`) — you do not need to add `FileType` unless targeting a specific format.

---

## Error Messages

| Message | Cause | Fix |
|---|---|---|
| `AI API key is not configured.` | No API key in AI Settings | Add a key in the AI Assistant tab → AI Settings |
| `Not connected. Configure credentials in the Authentication tab.` | OneDrive auth not set up or expired | Complete or refresh the Authentication tab setup |
| `Could not determine Graph API call. Please try rephrasing your question.` | AI returned a blank response | Rephrase with more specific terms |
| `Could not parse a valid Graph API call from the AI response.` | AI returned unparseable JSON | Rephrase the question more specifically |
| `Could not find the Excel file mentioned in your question.` | Filename not found in OneDrive search | Check the filename and extension; ensure the file is in your OneDrive |
| `Graph API call failed (HTTP 400): ...` | Invalid worksheet name, cell address, or body format | Check the sheet name, range address, and filename |
| `Graph API call failed (HTTP 403): ...` | Insufficient permissions on the Azure AD application | Add `Files.ReadWrite` scope (see [Authentication](pages/Authentication.md)) |
| `Graph API call failed: The template variable 'itemId' has no value` | File not found before `{itemId}` substitution | This error was fixed — include the exact filename with `.xlsx` extension in your question |
| `Query failed: ...` | Unexpected exception | Check extension logs for the underlying error |

---

## Troubleshooting

### GENERATED KQL shows the original question instead of a Graph API path

This was a bug in earlier versions where the generated plan was not passed through to the UI. It is fixed — the GENERATED KQL field now always shows the actual `method + path` that was executed (e.g. `PATCH /me/drive/items/ABC123/workbook/worksheets/Sheet1/range(address='A1')`).

### Write operation returns columns like `@odata.context` or `@odata.id`

These are internal Graph API metadata fields. They were visible in earlier versions and have been filtered out — PATCH write operations now return the friendly `"Cell updated successfully."` message instead.

### Cell read returns no `Value` column

If the cell is empty, the `values` field is `[[""]]` and no value is shown. If the range has data but the Value column is missing, ensure the question specifies the correct sheet name and range:

> "Get the value of A1 in Sheet1 in Budget.xlsx"

### Writing fails with a 400 error

The most common causes:

| Cause | Fix |
|---|---|
| Sheet name is wrong | Use the exact sheet name as it appears in Excel (case-sensitive) |
| Range address is wrong | Use standard Excel notation: `A1`, `A1:B5`, not `A 1` or `col1` |
| Values body is malformed | The body must be `{"values":[["value"]]}` — a 2D array |

### File found but wrong file is updated

When multiple files share the same name, the first search result is used. To target a specific file, include the folder path or a distinguishing part of the name in your question:

> "Update cell A1 in Sheet1 in the Budget.xlsx file in the Finance folder"

### `usedRange` returns too many columns

`usedRange` returns all used cells in the sheet. Noisy metadata fields (`numberFormat`, `formulas`, `text`, etc.) are filtered out automatically, but large sheets may still return many rows. For a specific area, use a range address instead:

> "Read cells A1 to D20 in Sheet1 in Budget.xlsx"

---

## Architecture Notes

- Both the NQL tab and Krista Agent tool calls use `ExcelMcpQueryService.query()` — the same pipeline.
- The AI generates `{method, path, body}` and the extension executes it via a JAX-RS HTTP client.
- All four HTTP methods are supported: `GET`, `POST`, `PATCH`, `DELETE`.
- `{itemId}` in AI-generated paths is auto-resolved via a OneDrive drive search by filename.
- `PATCH` responses return a friendly `"Cell updated successfully. (address)"` message instead of raw range metadata.
- Range metadata fields (formulas, numberFormat, hidden flags, etc.) are automatically filtered from results.
- Cell `values` arrays are flattened to readable strings (columns separated by ` \| `, rows by ` / `).
- Authentication credentials are reloaded on every call so changes to the Authentication tab take effect immediately without a restart.

---

## See Also

- [Authentication](pages/Authentication.md) — configure the OneDrive/Excel connection
- [Get Cell Value](pages/GetCellValue.md) — catalog request for reading a cell value
- [Add Data to Cell](pages/AddDataToCell.md) — catalog request for writing to a cell
- [Get Sheets Information](pages/GetSheetsInformation.md) — catalog request for listing worksheets
- [Get Table Rows](pages/GetTableRows.md) — catalog request for reading table rows
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md) — catalog request for listing shared Excel files
