/*
 * Office365excel 3.0 Extension for Krista
 * Copyright (C) 2025 Krista Software
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

package app.krista.extensions.essentials.spreadsheets.office365excel.mcp;

import app.krista.model.base.FreeForm;
import app.krista.sdk.msauth.extension.mcp.McpSearchResult;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ExcelMcpProductConfig}.
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>Product metadata and column definitions</li>
 *   <li>{@code buildSearchRequestBody} — correct Graph Search JSON structure</li>
 *   <li>{@code parseSearchResponse} — extracts rows and hitId from hitsContainers</li>
 *   <li>{@code parseStandardResponse} — maps FreeForm rows from hits</li>
 *   <li>{@code queryGenerationPrompt} — KQL syntax rules and examples</li>
 *   <li>{@code graphApiPlanningPrompt} — Graph API patterns for Excel/OneDrive</li>
 *   <li>{@code answerExtractionPrompt} — includes question and content data</li>
 * </ul>
 */
class ExcelMcpProductConfigTest {

    private ExcelMcpProductConfig config;

    @BeforeEach
    void setUp() {
        config = new ExcelMcpProductConfig();
    }

    // ── Product metadata ──────────────────────────────────────────────────

    @Test
    void productName_isOffice365Excel() {
        assertEquals("Office 365 Excel", config.productName());
    }

    @Test
    void queryLanguage_isKql() {
        assertEquals("KQL", config.queryLanguage());
    }

    @Test
    void queryLanguageFullName_isKqlFull() {
        assertEquals("KQL (Keyword Query Language)", config.queryLanguageFullName());
    }

    @Test
    void searchHttpMethod_isPost() {
        assertEquals("POST", config.searchHttpMethod());
    }

    @Test
    void searchUrl_isGraphSearchQuery() {
        assertTrue(config.searchUrl().contains("graph.microsoft.com"),
                "searchUrl must point to Microsoft Graph");
        assertTrue(config.searchUrl().contains("search/query"),
                "searchUrl must use the /search/query endpoint");
    }

    // ── resultColumns returns empty (dynamic columns) ─────────────────────

    @Test
    void resultColumns_isEmpty() {
        assertTrue(config.resultColumns().isEmpty(),
                "resultColumns must be empty — columns are derived dynamically at runtime");
    }

    // ── standardFields ────────────────────────────────────────────────────

    @Test
    void standardFields_containsEssentialFileFields() {
        List<String> fields = config.standardFields();
        assertTrue(fields.contains("name"),                 "must include name");
        assertTrue(fields.contains("webUrl"),               "must include webUrl");
        assertTrue(fields.contains("lastModifiedDateTime"), "must include lastModifiedDateTime");
        assertTrue(fields.contains("lastModifiedBy"),       "must include lastModifiedBy");
        assertTrue(fields.contains("size"),                 "must include size");
    }

    // ── nlExamples and kqlExamples ────────────────────────────────────────

    @Test
    void nlExamples_notEmpty() {
        assertFalse(config.nlExamples().isEmpty(),
                "nlExamples must provide at least one example for the UI");
    }

    @Test
    void kqlExamples_notEmpty() {
        assertFalse(config.kqlExamples().isEmpty(),
                "kqlExamples must provide at least one KQL example");
    }

    // ── buildSearchRequestBody ────────────────────────────────────────────

    @Test
    void buildSearchRequestBody_producesValidJson() {
        String body = config.buildSearchRequestBody("budget", 10, false);
        assertDoesNotThrow(() -> JsonParser.parseString(body),
                "buildSearchRequestBody must produce valid JSON");
    }

    @Test
    void buildSearchRequestBody_entityTypeIsDriveItem() {
        String body = config.buildSearchRequestBody("*", 10, false);
        assertTrue(body.contains("\"driveItem\""),
                "entityTypes must include 'driveItem' for OneDrive/SharePoint file search");
    }

    @Test
    void buildSearchRequestBody_containsQueryString() {
        String body = config.buildSearchRequestBody("budget", 5, false);
        JsonObject root = JsonParser.parseString(body).getAsJsonObject();
        JsonObject request = root.getAsJsonArray("requests").get(0).getAsJsonObject();
        String queryString = request.getAsJsonObject("query").get("queryString").getAsString();
        assertTrue(queryString.contains("budget"), "queryString must include the supplied KQL query");
    }

    @Test
    void buildSearchRequestBody_respectsMaxResults() {
        String body = config.buildSearchRequestBody("*", 25, false);
        JsonObject root = JsonParser.parseString(body).getAsJsonObject();
        JsonObject request = root.getAsJsonArray("requests").get(0).getAsJsonObject();
        assertEquals(25, request.get("size").getAsInt());
    }

    @Test
    void buildSearchRequestBody_queryWithoutFileType_appendsExcelFilter() {
        String body = config.buildSearchRequestBody("budget", 10, false);
        assertTrue(body.contains("FileType:xlsx"),
                "When no FileType is already present, Excel file type filter must be appended");
    }

    @Test
    void buildSearchRequestBody_queryWithFileType_doesNotDuplicate() {
        String body = config.buildSearchRequestBody("FileType:xlsx", 10, false);
        int count = body.split("FileType:", -1).length - 1;
        assertEquals(1, count,
                "FileType filter must not be duplicated when already present in the query");
    }

    // ── parseSearchResponse ───────────────────────────────────────────────

    @Test
    void parseSearchResponse_emptyHits_returnsZeroRows() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":0,\"hits\":[]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals(0, result.shown());
    }

    @Test
    void parseSearchResponse_singleFile_extractsName() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "  \"resource\":{"
                + "    \"name\":\"Budget.xlsx\","
                + "    \"webUrl\":\"https://onedrive.com/1\","
                + "    \"lastModifiedDateTime\":\"2024-06-01T09:00:00Z\","
                + "    \"size\":\"2048\""
                + "  }"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals(1, result.shown());
        assertEquals("Budget.xlsx", result.rows().get(0).get("Name"),
                "File name must be mapped to 'Name' key");
    }

    @Test
    void parseSearchResponse_extractsHitIdAsFileId() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "  \"hitId\":\"file-abc-123\","
                + "  \"resource\":{\"name\":\"Test.xlsx\"}"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("file-abc-123", result.rows().get(0).get("fileId"),
                "hitId from the hit object must be stored as 'fileId'");
    }

    @Test
    void parseSearchResponse_truncatesDateTimeToDate() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "  \"resource\":{"
                + "    \"name\":\"File.xlsx\","
                + "    \"lastModifiedDateTime\":\"2024-05-15T10:30:00Z\""
                + "  }"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        String lastModified = (String) result.rows().get(0).get("LastModified");
        assertNotNull(lastModified);
        assertFalse(lastModified.contains("T"),
                "lastModifiedDateTime must be truncated to date only (no time portion)");
    }

    @Test
    void parseSearchResponse_formatsFileSizeInKb() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "  \"resource\":{"
                + "    \"name\":\"File.xlsx\","
                + "    \"size\":\"2048\""
                + "  }"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("2 KB", result.rows().get(0).get("Size"),
                "File size in bytes must be formatted as KB");
    }

    @Test
    void parseSearchResponse_multipleHits_returnsAllRows() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"total\":2,\"hits\":["
                + "  {\"resource\":{\"name\":\"File1.xlsx\"}},"
                + "  {\"resource\":{\"name\":\"File2.xlsx\"}}"
                + "]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals(2, result.shown());
    }

    // ── parseStandardResponse ─────────────────────────────────────────────

    @Test
    void parseStandardResponse_emptyHits_returnsEmptyList() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertTrue(rows.isEmpty());
    }

    @Test
    void parseStandardResponse_singleFile_returnsFreeFormWithName() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "  \"resource\":{\"name\":\"Budget.xlsx\",\"webUrl\":\"https://onedrive.com/1\"}"
                + "}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertEquals(1, rows.size());
        assertNotNull(rows.get(0).get("Name"), "FreeForm must contain 'Name' key");
    }

    @Test
    void parseStandardResponse_extractsHitIdFromHit() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "  \"hitId\":\"hit-xyz-789\","
                + "  \"resource\":{\"name\":\"Workbook.xlsx\"}"
                + "}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertNotNull(rows.get(0).get("FileId"),
                "FileId FreeForm key must be populated from the hit's hitId");
    }

    @Test
    void parseStandardResponse_multipleFiles_returnsAll() {
        String json = "{"
                + "\"value\":[{\"hitsContainers\":[{\"hits\":["
                + "  {\"resource\":{\"name\":\"File1.xlsx\"}},"
                + "  {\"resource\":{\"name\":\"File2.xlsx\"}}"
                + "]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertEquals(2, rows.size());
    }

    // ── queryGenerationPrompt ─────────────────────────────────────────────

    @Test
    void queryGenerationPrompt_containsQuestion() {
        String prompt = config.queryGenerationPrompt("budget spreadsheets");
        assertTrue(prompt.contains("budget spreadsheets"));
    }

    @Test
    void queryGenerationPrompt_containsKqlProperties() {
        String prompt = config.queryGenerationPrompt("test");
        assertTrue(prompt.contains("FileType"),          "must mention FileType property");
        assertTrue(prompt.contains("Author"),            "must mention Author property");
        assertTrue(prompt.contains("LastModifiedTime"),  "must mention LastModifiedTime");
        assertTrue(prompt.contains("Title"),             "must mention Title property");
    }

    @Test
    void queryGenerationPrompt_includesExamples() {
        String prompt = config.queryGenerationPrompt("test");
        assertTrue(prompt.contains("Author:"),              "must include Author example");
        assertTrue(prompt.contains("LastModifiedTime>"),    "must include date range example");
    }

    @Test
    void queryGenerationPrompt_instructsOutputOnlyJson() {
        String prompt = config.queryGenerationPrompt("test");
        assertTrue(prompt.contains("Output ONLY raw JSON"));
    }

    @Test
    void queryGenerationPrompt_outputSchemaContainsQueryField() {
        String prompt = config.queryGenerationPrompt("test");
        assertTrue(prompt.contains("\"query\""),
                "Output schema must include the 'query' field");
    }

    @Test
    void queryGenerationPrompt_mentionsExcelFileTypes() {
        String prompt = config.queryGenerationPrompt("test");
        assertTrue(prompt.contains("xlsx"),  "must mention xlsx file type");
        assertTrue(prompt.contains("xls"),   "must mention xls file type");
        assertTrue(prompt.contains("xlsm"),  "must mention xlsm file type");
    }

    // ── graphApiPlanningPrompt ────────────────────────────────────────────

    @Test
    void graphApiPlanningPrompt_containsQuestion() {
        String prompt = config.graphApiPlanningPrompt("show my recent Excel files");
        assertTrue(prompt.contains("show my recent Excel files"));
    }

    @Test
    void graphApiPlanningPrompt_containsSearchQueryPattern() {
        String prompt = config.graphApiPlanningPrompt("test");
        assertTrue(prompt.contains("/search/query"),
                "Prompt must include search/query pattern for file search");
    }

    @Test
    void graphApiPlanningPrompt_containsDrivePatterns() {
        String prompt = config.graphApiPlanningPrompt("test");
        assertTrue(prompt.contains("/me/drive/root/children"),
                "Prompt must include drive root children pattern");
        assertTrue(prompt.contains("/me/drive/recent"),
                "Prompt must include recent files pattern");
    }

    @Test
    void graphApiPlanningPrompt_containsWorkbookPatterns() {
        String prompt = config.graphApiPlanningPrompt("test");
        assertTrue(prompt.contains("/workbook/worksheets"),
                "Prompt must include worksheets pattern for workbook operations");
    }

    @Test
    void graphApiPlanningPrompt_containsWritePattern() {
        String prompt = config.graphApiPlanningPrompt("test");
        assertTrue(prompt.contains("PATCH"),
                "Prompt must include PATCH method for write cell operations");
    }

    @Test
    void graphApiPlanningPrompt_outputIsRawJsonOnly() {
        String prompt = config.graphApiPlanningPrompt("test");
        assertTrue(prompt.contains("Output ONLY a raw JSON object"),
                "Prompt must instruct AI to output raw JSON only");
    }

    @Test
    void graphApiPlanningPrompt_outputSchemaContainsMethodPathBody() {
        String prompt = config.graphApiPlanningPrompt("test");
        assertTrue(prompt.contains("\"method\""), "Output schema must include method");
        assertTrue(prompt.contains("\"path\""),   "Output schema must include path");
        assertTrue(prompt.contains("\"body\""),   "Output schema must include body");
    }

    @Test
    void graphApiPlanningPrompt_mentionsItemIdPlaceholder() {
        String prompt = config.graphApiPlanningPrompt("test");
        assertTrue(prompt.contains("{itemId}"),
                "Prompt must mention {itemId} placeholder for file ID auto-resolution");
    }

    // ── answerExtractionPrompt ────────────────────────────────────────────

    @Test
    void answerExtractionPrompt_containsQuestion() {
        String prompt = config.answerExtractionPrompt("how many files do I have?", "[]");
        assertTrue(prompt.contains("how many files do I have?"));
    }

    @Test
    void answerExtractionPrompt_containsSearchResults() {
        String prompt = config.answerExtractionPrompt("q", "[{\"name\":\"Budget.xlsx\"}]");
        assertTrue(prompt.contains("[{\"name\":\"Budget.xlsx\"}]"));
    }

    @Test
    void answerExtractionPrompt_instructsNotToHallucinate() {
        String prompt = config.answerExtractionPrompt("q", "data");
        assertTrue(prompt.contains("Do not invent") || prompt.contains("hallucinate"),
                "Prompt must instruct AI not to hallucinate data");
    }
}
