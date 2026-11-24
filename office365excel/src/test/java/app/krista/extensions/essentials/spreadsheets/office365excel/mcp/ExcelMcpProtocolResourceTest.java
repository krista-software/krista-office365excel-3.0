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

import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.sdk.msauth.token.ITokenStore;
import com.google.gson.Gson;
import com.google.gson.JsonNull;
import com.google.gson.JsonObject;
import com.google.gson.JsonPrimitive;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import javax.ws.rs.core.Response;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.*;

/**
 * Unit tests for {@link ExcelMcpProtocolResource}.
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>{@code handle()} — all dispatch branches (parse error, initialize, ping,
 *       tools/list, notifications/initialized, unknown method, tools/call)</li>
 *   <li>{@code buildInitializeResult} — via handle()</li>
 *   <li>{@code buildToolsList} — via handle()</li>
 *   <li>{@code formatSearchResponse} — all branches via reflection</li>
 *   <li>{@code formatSingleItem} — all branches via reflection</li>
 *   <li>{@code formatWorksheetsResponse} — all branches via reflection</li>
 *   <li>{@code formatWorksheetData} — all branches via reflection</li>
 *   <li>{@code appendRow / appendSeparator / pad} — via reflection</li>
 *   <li>{@code wrapGraphError} — via reflection</li>
 *   <li>{@code extractDisplayName / truncateDate / formatSize / getStr / getObj} — via reflection</li>
 *   <li>{@code buildSearchBody / buildTool / buildSchema / strProp} — via reflection</li>
 *   <li>{@code ok / toolResult / toolError / mcpError} — via reflection</li>
 *   <li>{@code resolvedPath} — via reflection</li>
 *   <li>{@code adaptTokenStore} — save/get/delete and IOException wrapping</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class ExcelMcpProtocolResourceTest {

    @Mock private RefreshTokenStore         refreshTokenStore;
    @Mock private Invoker                   invoker;
    @Mock private OfficeExcelAttributeStore attributeStore;

    private ExcelMcpProtocolResource resource;
    private static final Gson GSON = new Gson();

    @BeforeEach
    void setUp() throws Exception {
        lenient().when(invoker.getInvokerId()).thenReturn("test-invoker");
        lenient().when(attributeStore.load("test-invoker")).thenReturn(null);
        lenient().when(refreshTokenStore.get(anyString())).thenReturn(null);
        resource = new ExcelMcpProtocolResource(refreshTokenStore, invoker, attributeStore);
    }

    // ── handle() — dispatcher ────────────────────────────────────────────────

    @Test
    void handle_invalidJson_returnsParseError() {
        Response resp = resource.handle("not-json{{{");
        JsonObject json = parseEntity(resp);
        assertEquals(-32700, json.getAsJsonObject("error").get("code").getAsInt());
    }

    @Test
    void handle_invalidJson_errorMessageContainsParseError() {
        Response resp = resource.handle("{bad json");
        JsonObject json = parseEntity(resp);
        assertTrue(json.getAsJsonObject("error").get("message").getAsString()
                .startsWith("Parse error"));
    }

    @Test
    void handle_initialize_containsProtocolVersion() {
        JsonObject result = invokeHandle("initialize");
        assertEquals("2024-11-05", result.get("protocolVersion").getAsString());
    }

    @Test
    void handle_initialize_serverNameIsKristaExcelMcp() {
        JsonObject result = invokeHandle("initialize");
        assertEquals("krista-excel-mcp",
                result.getAsJsonObject("serverInfo").get("name").getAsString());
    }

    @Test
    void handle_initialize_serverVersionIs1_0_0() {
        JsonObject result = invokeHandle("initialize");
        assertEquals("1.0.0",
                result.getAsJsonObject("serverInfo").get("version").getAsString());
    }

    @Test
    void handle_initialize_capabilitiesPresent() {
        JsonObject result = invokeHandle("initialize");
        assertNotNull(result.get("capabilities"));
    }

    @Test
    void handle_notificationsInitialized_returnsEmptyObject() {
        String req = "{\"jsonrpc\":\"2.0\",\"method\":\"notifications/initialized\"}";
        Response resp = resource.handle(req);
        assertEquals("{}", resp.getEntity());
    }

    @Test
    void handle_ping_returnsEmptyResult() {
        JsonObject result = invokeHandle("ping");
        assertTrue(result.entrySet().isEmpty());
    }

    @Test
    void handle_toolsList_returnsFiveTools() {
        JsonObject result = invokeHandle("tools/list");
        assertEquals(5, result.getAsJsonArray("tools").size());
    }

    @Test
    void handle_toolsList_containsListWorkbooks() {
        JsonObject result = invokeHandle("tools/list");
        assertTrue(toolNames(result).contains("list_workbooks"));
    }

    @Test
    void handle_toolsList_containsSearchWorkbooks() {
        JsonObject result = invokeHandle("tools/list");
        assertTrue(toolNames(result).contains("search_workbooks"));
    }

    @Test
    void handle_toolsList_containsGetWorkbook() {
        JsonObject result = invokeHandle("tools/list");
        assertTrue(toolNames(result).contains("get_workbook"));
    }

    @Test
    void handle_toolsList_containsListWorksheets() {
        JsonObject result = invokeHandle("tools/list");
        assertTrue(toolNames(result).contains("list_worksheets"));
    }

    @Test
    void handle_toolsList_containsReadWorksheet() {
        JsonObject result = invokeHandle("tools/list");
        assertTrue(toolNames(result).contains("read_worksheet"));
    }

    @Test
    void handle_toolsList_eachToolHasInputSchema() {
        JsonObject result = invokeHandle("tools/list");
        result.getAsJsonArray("tools").forEach(t ->
                assertNotNull(t.getAsJsonObject().get("inputSchema"),
                        "Each tool must have an inputSchema"));
    }

    @Test
    void handle_unknownMethod_returnsMethodNotFound() {
        JsonObject json = parseEntity(resource.handle(
                "{\"jsonrpc\":\"2.0\",\"id\":4,\"method\":\"nonexistent/method\"}"));
        assertEquals(-32601, json.getAsJsonObject("error").get("code").getAsInt());
    }

    @Test
    void handle_emptyMethod_returnsMethodNotFound() {
        JsonObject json = parseEntity(resource.handle("{\"jsonrpc\":\"2.0\",\"id\":5}"));
        assertEquals(-32601, json.getAsJsonObject("error").get("code").getAsInt());
    }

    @Test
    void handle_noIdField_responseHasNoId() {
        // notification — no "id" → response must not include "id"
        JsonObject json = parseEntity(resource.handle(
                "{\"jsonrpc\":\"2.0\",\"method\":\"ping\"}"));
        assertFalse(json.has("id"));
    }

    @Test
    void handle_withNullId_responseHasNoId() {
        JsonObject json = parseEntity(resource.handle(
                "{\"jsonrpc\":\"2.0\",\"id\":null,\"method\":\"ping\"}"));
        assertFalse(json.has("id"), "Explicit null id must not be echoed back");
    }

    @Test
    void handle_withStringId_responseEchoesId() {
        JsonObject json = parseEntity(resource.handle(
                "{\"jsonrpc\":\"2.0\",\"id\":\"req-1\",\"method\":\"ping\"}"));
        assertEquals("req-1", json.get("id").getAsString());
    }

    // ── handle() — tools/call — not authenticated ────────────────────────────

    @Test
    void handle_toolsCall_notAuthenticated_isErrorTrue() {
        JsonObject result = toolsCallResult("list_workbooks", "{}");
        assertTrue(result.get("isError").getAsBoolean());
    }

    @Test
    void handle_toolsCall_notAuthenticated_textContainsNotConnected() {
        JsonObject result = toolsCallResult("list_workbooks", "{}");
        assertTrue(contentText(result).contains("Not connected"));
    }

    @Test
    void handle_toolsCall_internalException_returnsMcpError() {
        // Make attributeStore.load() throw so currentAuthBridge() propagates an exception
        when(attributeStore.load("test-invoker")).thenThrow(new RuntimeException("db failure"));
        JsonObject json = parseEntity(resource.handle(
                "{\"jsonrpc\":\"2.0\",\"id\":7,\"method\":\"tools/call\","
                        + "\"params\":{\"name\":\"list_workbooks\",\"arguments\":{}}}"));
        assertEquals(-32603, json.getAsJsonObject("error").get("code").getAsInt());
    }

    @Test
    void handle_toolsCall_unknownTool_notAuthenticated_isErrorTrue() {
        // Unknown tool is only reached after auth; with null attrs auth always fails
        // → we get the "not authenticated" toolError regardless of tool name
        JsonObject result = toolsCallResult("unknown_tool_xyz", "{}");
        assertTrue(result.get("isError").getAsBoolean());
    }

    // ── formatSearchResponse ─────────────────────────────────────────────────

    @Test
    void formatSearchResponse_emptyValueArray_returnsNoLabel() throws Exception {
        assertEquals("No workbooks found.",
                invokeFormatSearchResponse("{\"value\":[]}", "workbooks"));
    }

    @Test
    void formatSearchResponse_graphError_returnsGraphErrorMessage() throws Exception {
        String result = invokeFormatSearchResponse(
                "{\"error\":{\"message\":\"Access denied\"}}", "workbooks");
        assertTrue(result.contains("Access denied"));
    }

    @Test
    void formatSearchResponse_withHit_containsFileName() throws Exception {
        String result = invokeFormatSearchResponse(searchJson("Budget.xlsx", "abc-123"), "workbooks");
        assertTrue(result.contains("Budget.xlsx"));
    }

    @Test
    void formatSearchResponse_withHit_containsTruncatedDate() throws Exception {
        String result = invokeFormatSearchResponse(searchJson("Budget.xlsx", "abc-123"), "workbooks");
        assertTrue(result.contains("2024-05-15"));
    }

    @Test
    void formatSearchResponse_withHit_containsFormattedSize() throws Exception {
        String result = invokeFormatSearchResponse(searchJson("Budget.xlsx", "abc-123"), "workbooks");
        assertTrue(result.contains("2 KB"));
    }

    @Test
    void formatSearchResponse_withHit_containsFileId() throws Exception {
        String result = invokeFormatSearchResponse(searchJson("Budget.xlsx", "abc-123"), "workbooks");
        assertTrue(result.contains("abc-123"));
    }

    @Test
    void formatSearchResponse_withLastModifiedBy_showsDisplayName() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":\"id1\","
                + "\"resource\":{"
                + "  \"name\":\"Report.xlsx\","
                + "  \"lastModifiedBy\":{\"user\":{\"displayName\":\"Alice\"}}"
                + "}}]}]}]}";
        String result = invokeFormatSearchResponse(json, "workbooks");
        assertTrue(result.contains("Alice"));
    }

    @Test
    void formatSearchResponse_resourceIdFallback_whenNoHitId() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"id\":\"resource-id-999\","
                + "  \"name\":\"Test.xlsx\""
                + "}}]}]}]}";
        String result = invokeFormatSearchResponse(json, "workbooks");
        assertTrue(result.contains("resource-id-999"));
    }

    @Test
    void formatSearchResponse_multipleHits_countInHeader() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":["
                + "{\"hitId\":\"1\",\"resource\":{\"name\":\"A.xlsx\"}},"
                + "{\"hitId\":\"2\",\"resource\":{\"name\":\"B.xlsx\"}}"
                + "]}]}]}";
        String result = invokeFormatSearchResponse(json, "workbooks");
        assertTrue(result.startsWith("Found 2 workbooks"));
    }

    @Test
    void formatSearchResponse_emptyHitsContainer_returnsNoLabel() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[]}]}]}";
        String result = invokeFormatSearchResponse(json, "files");
        assertEquals("No files found.", result);
    }

    @Test
    void formatSearchResponse_noHitsContainersKey_returnsNoLabel() throws Exception {
        String json = "{\"value\":[{}]}";
        String result = invokeFormatSearchResponse(json, "docs");
        assertEquals("No docs found.", result);
    }

    @Test
    void formatSearchResponse_webUrlPresent_showsLink() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":\"h1\","
                + "\"resource\":{\"name\":\"F.xlsx\",\"webUrl\":\"https://sharepoint.com/f\"}"
                + "}]}]}]}";
        String result = invokeFormatSearchResponse(json, "workbooks");
        assertTrue(result.contains("https://sharepoint.com/f"));
    }

    // ── formatSingleItem ─────────────────────────────────────────────────────

    @Test
    void formatSingleItem_normalObject_containsName() throws Exception {
        assertTrue(invokeFormatSingleItem(singleItemJson()).contains("Budget.xlsx"));
    }

    @Test
    void formatSingleItem_normalObject_containsId() throws Exception {
        assertTrue(invokeFormatSingleItem(singleItemJson()).contains("abc-123"));
    }

    @Test
    void formatSingleItem_normalObject_containsFormattedSize() throws Exception {
        assertTrue(invokeFormatSingleItem(singleItemJson()).contains("1.0 MB"));
    }

    @Test
    void formatSingleItem_normalObject_containsTruncatedDate() throws Exception {
        assertTrue(invokeFormatSingleItem(singleItemJson()).contains("2024-05-15"));
    }

    @Test
    void formatSingleItem_normalObject_containsWebUrl() throws Exception {
        assertTrue(invokeFormatSingleItem(singleItemJson()).contains("https://onedrive.com/file"));
    }

    @Test
    void formatSingleItem_graphError_returnsErrorMessage() throws Exception {
        String result = invokeFormatSingleItem("{\"error\":{\"message\":\"Item not found\"}}");
        assertTrue(result.contains("Item not found"));
    }

    @Test
    void formatSingleItem_withLastModifiedBy_showsDisplayName() throws Exception {
        String json = "{\"name\":\"File.xlsx\",\"id\":\"x1\","
                + "\"lastModifiedBy\":{\"user\":{\"displayName\":\"Bob\"}}}";
        assertTrue(invokeFormatSingleItem(json).contains("Bob"));
    }

    @Test
    void formatSingleItem_noSizeField_noSizeLine() throws Exception {
        String result = invokeFormatSingleItem("{\"name\":\"Minimal.xlsx\",\"id\":\"m1\"}");
        assertFalse(result.contains("Size:"));
    }

    @Test
    void formatSingleItem_noModifiedField_noModifiedLine() throws Exception {
        String result = invokeFormatSingleItem("{\"name\":\"Minimal.xlsx\",\"id\":\"m1\"}");
        assertFalse(result.contains("Modified:"));
    }

    @Test
    void formatSingleItem_noWebUrl_noLinkLine() throws Exception {
        String result = invokeFormatSingleItem("{\"name\":\"Minimal.xlsx\",\"id\":\"m1\"}");
        assertFalse(result.contains("Link:"));
    }

    // ── formatWorksheetsResponse ─────────────────────────────────────────────

    @Test
    void formatWorksheetsResponse_graphError_returnsError() throws Exception {
        String result = invokeFormatWorksheetsResponse("{\"error\":{\"message\":\"File not found\"}}");
        assertTrue(result.contains("File not found"));
    }

    @Test
    void formatWorksheetsResponse_emptyValue_returnsNoWorksheetsMessage() throws Exception {
        assertEquals("No worksheets found in this workbook.",
                invokeFormatWorksheetsResponse("{\"value\":[]}"));
    }

    @Test
    void formatWorksheetsResponse_singleVisibleSheet_noVisibilityLabel() throws Exception {
        String result = invokeFormatWorksheetsResponse(
                "{\"value\":[{\"name\":\"Sheet1\",\"visibility\":\"Visible\"}]}");
        assertTrue(result.contains("Sheet1"));
        assertFalse(result.contains("(visible)"));
    }

    @Test
    void formatWorksheetsResponse_hiddenSheet_showsHiddenLabel() throws Exception {
        String result = invokeFormatWorksheetsResponse(
                "{\"value\":[{\"name\":\"HiddenTab\",\"visibility\":\"Hidden\"}]}");
        assertTrue(result.contains("(hidden)"));
    }

    @Test
    void formatWorksheetsResponse_twoSheets_countInHeader() throws Exception {
        String json = "{\"value\":["
                + "{\"name\":\"Sheet1\",\"visibility\":\"Visible\"},"
                + "{\"name\":\"Sheet2\",\"visibility\":\"Visible\"}"
                + "]}";
        assertTrue(invokeFormatWorksheetsResponse(json).startsWith("2 worksheet(s)"));
    }

    @Test
    void formatWorksheetsResponse_veryHiddenSheet_showsLabel() throws Exception {
        String result = invokeFormatWorksheetsResponse(
                "{\"value\":[{\"name\":\"Tab\",\"visibility\":\"VeryHidden\"}]}");
        assertTrue(result.contains("(veryhidden)"));
    }

    // ── formatWorksheetData ──────────────────────────────────────────────────

    @Test
    void formatWorksheetData_graphError_returnsError() throws Exception {
        String result = invokeFormatWorksheetData("{\"error\":{\"message\":\"Unauthorized\"}}", "Sheet1");
        assertTrue(result.contains("Unauthorized"));
    }

    @Test
    void formatWorksheetData_emptyValues_returnsEmptyMessage() throws Exception {
        assertEquals("Sheet \"Sheet1\" is empty.",
                invokeFormatWorksheetData("{\"values\":[]}", "Sheet1"));
    }

    @Test
    void formatWorksheetData_twoByTwo_containsHeader() throws Exception {
        String result = invokeFormatWorksheetData(
                "{\"values\":[[\"Name\",\"Age\"],[\"Alice\",\"30\"]]}", "Data");
        assertTrue(result.contains("Name"));
        assertTrue(result.contains("Age"));
    }

    @Test
    void formatWorksheetData_twoByTwo_containsDataRow() throws Exception {
        String result = invokeFormatWorksheetData(
                "{\"values\":[[\"Name\",\"Age\"],[\"Alice\",\"30\"]]}", "Data");
        assertTrue(result.contains("Alice"));
        assertTrue(result.contains("30"));
    }

    @Test
    void formatWorksheetData_twoByTwo_containsSeparatorLine() throws Exception {
        String result = invokeFormatWorksheetData(
                "{\"values\":[[\"Name\",\"Age\"],[\"Alice\",\"30\"]]}", "Data");
        assertTrue(result.contains("---"));
    }

    @Test
    void formatWorksheetData_sheetNameInHeader() throws Exception {
        String result = invokeFormatWorksheetData(
                "{\"values\":[[\"Col\"],[\"Val\"]]}", "MySheet");
        assertTrue(result.contains("MySheet"));
    }

    @Test
    void formatWorksheetData_moreThan101Rows_showsEllipsis() throws Exception {
        StringBuilder rows = new StringBuilder("[\"Header\"]");
        for (int i = 0; i < 104; i++) rows.append(",[\"Row").append(i).append("\"]");
        String result = invokeFormatWorksheetData("{\"values\":[" + rows + "]}", "Big");
        assertTrue(result.contains("more rows"));
    }

    @Test
    void formatWorksheetData_cellLongerThan30Chars_isTruncated() throws Exception {
        String longVal = "A".repeat(35);
        String result = invokeFormatWorksheetData(
                "{\"values\":[[\"" + longVal + "\"],[\"x\"]]}", "Sheet1");
        assertTrue(result.contains("..."));
    }

    @Test
    void formatWorksheetData_rowWith101Rows_neitherEllipsisNorTruncation() throws Exception {
        // exactly 101 rows — limit = min(101, 101) = 101, loop runs r=1..100 (100 data rows),
        // table.size() == 101 which is NOT > 101 → no ellipsis
        StringBuilder rows = new StringBuilder("[\"Header\"]");
        for (int i = 0; i < 100; i++) rows.append(",[\"Row").append(i).append("\"]");
        String result = invokeFormatWorksheetData("{\"values\":[" + rows + "]}", "Sheet");
        assertFalse(result.contains("more rows"));
    }

    // ── appendRow ────────────────────────────────────────────────────────────

    @Test
    void appendRow_twoColumns_appendsSeparatorBetween() throws Exception {
        StringBuilder sb = new StringBuilder();
        invokeAppendRow(sb, List.of("A", "B"), new int[]{3, 3});
        assertTrue(sb.toString().contains("|"));
    }

    @Test
    void appendRow_singleColumn_noSeparator() throws Exception {
        StringBuilder sb = new StringBuilder();
        invokeAppendRow(sb, List.of("A"), new int[]{3});
        assertFalse(sb.toString().contains("|"));
    }

    @Test
    void appendRow_emptyRowList_filledWithEmpty() throws Exception {
        StringBuilder sb = new StringBuilder();
        // row shorter than widths — empty fallback
        invokeAppendRow(sb, List.of(), new int[]{3, 3});
        String result = sb.toString();
        assertTrue(result.contains("|"));
    }

    // ── appendSeparator ──────────────────────────────────────────────────────

    @Test
    void appendSeparator_twoColumns_containsDashPlusDash() throws Exception {
        StringBuilder sb = new StringBuilder();
        invokeAppendSeparator(sb, new int[]{3, 3});
        assertTrue(sb.toString().contains("-+-"));
    }

    @Test
    void appendSeparator_singleColumn_noPlusSeparator() throws Exception {
        StringBuilder sb = new StringBuilder();
        invokeAppendSeparator(sb, new int[]{5});
        assertFalse(sb.toString().contains("+"));
    }

    // ── pad ──────────────────────────────────────────────────────────────────

    @Test
    void pad_shorterString_padsToWidth() throws Exception {
        assertEquals("Hi   ", invokePad("Hi", 5));
    }

    @Test
    void pad_exactLength_unchanged() throws Exception {
        assertEquals("Hello", invokePad("Hello", 5));
    }

    @Test
    void pad_longerString_returnedUnchanged() throws Exception {
        assertEquals("TooLong", invokePad("TooLong", 3));
    }

    // ── wrapGraphError ───────────────────────────────────────────────────────

    @Test
    void wrapGraphError_bodyAlreadyHasErrorKey_passesThroughUnchanged() throws Exception {
        String body = "{\"error\":{\"code\":\"AccessDenied\",\"message\":\"Access denied\"}}";
        String result = invokeWrapGraphError(403, body);
        assertTrue(result.contains("Access denied"));
    }

    @Test
    void wrapGraphError_plainBody_wrapsWithStatusCode() throws Exception {
        JsonObject wrapped = GSON.fromJson(invokeWrapGraphError(500, "server error"), JsonObject.class);
        assertTrue(wrapped.has("error"));
        assertEquals("500", wrapped.getAsJsonObject("error").get("code").getAsString());
    }

    @Test
    void wrapGraphError_nonJsonBody_wrapsBody() throws Exception {
        JsonObject wrapped = GSON.fromJson(invokeWrapGraphError(401, "Unauthorized text"), JsonObject.class);
        assertTrue(wrapped.getAsJsonObject("error").get("message").getAsString().contains("401"));
    }

    // ── extractDisplayName ───────────────────────────────────────────────────

    @Test
    void extractDisplayName_withUser_returnsDisplayName() throws Exception {
        JsonObject parent = buildIdentitySet("lastModifiedBy", "Alice Smith");
        assertEquals("Alice Smith", invokeExtractDisplayName(parent, "lastModifiedBy"));
    }

    @Test
    void extractDisplayName_missingField_returnsEmpty() throws Exception {
        assertEquals("", invokeExtractDisplayName(new JsonObject(), "lastModifiedBy"));
    }

    @Test
    void extractDisplayName_nullField_returnsEmpty() throws Exception {
        JsonObject parent = new JsonObject();
        parent.add("lastModifiedBy", JsonNull.INSTANCE);
        assertEquals("", invokeExtractDisplayName(parent, "lastModifiedBy"));
    }

    @Test
    void extractDisplayName_missingUserKey_returnsEmpty() throws Exception {
        JsonObject parent = new JsonObject();
        parent.add("lastModifiedBy", new JsonObject());
        assertEquals("", invokeExtractDisplayName(parent, "lastModifiedBy"));
    }

    @Test
    void extractDisplayName_nullUser_returnsEmpty() throws Exception {
        JsonObject parent = new JsonObject();
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", JsonNull.INSTANCE);
        parent.add("lastModifiedBy", identitySet);
        assertEquals("", invokeExtractDisplayName(parent, "lastModifiedBy"));
    }

    // ── truncateDate ─────────────────────────────────────────────────────────

    @Test
    void truncateDate_withTSuffix_returnsDateOnly() throws Exception {
        assertEquals("2024-05-15", invokeTruncateDate("2024-05-15T10:30:00Z"));
    }

    @Test
    void truncateDate_noT_returnsAsIs() throws Exception {
        assertEquals("2024-05-15", invokeTruncateDate("2024-05-15"));
    }

    @Test
    void truncateDate_null_returnsEmpty() throws Exception {
        assertEquals("", invokeTruncateDate(null));
    }

    @Test
    void truncateDate_empty_returnsEmpty() throws Exception {
        assertEquals("", invokeTruncateDate(""));
    }

    // ── formatSize ───────────────────────────────────────────────────────────

    @Test
    void formatSize_null_returnsEmpty() throws Exception {
        assertEquals("", invokeFormatSize(null));
    }

    @Test
    void formatSize_empty_returnsEmpty() throws Exception {
        assertEquals("", invokeFormatSize(""));
    }

    @Test
    void formatSize_lessThan1024_returnsBytes() throws Exception {
        assertEquals("500 B", invokeFormatSize("500"));
    }

    @Test
    void formatSize_between1024And1MB_returnsKB() throws Exception {
        assertEquals("2 KB", invokeFormatSize("2048"));
    }

    @Test
    void formatSize_oneExactMB_returnsMB() throws Exception {
        assertEquals("1.0 MB", invokeFormatSize("1048576"));
    }

    @Test
    void formatSize_nonNumeric_returnsAsIs() throws Exception {
        assertEquals("unknown", invokeFormatSize("unknown"));
    }

    // ── getStr ───────────────────────────────────────────────────────────────

    @Test
    void getStr_nullObject_returnsEmpty() throws Exception {
        assertEquals("", invokeGetStr(null, "key"));
    }

    @Test
    void getStr_missingKey_returnsEmpty() throws Exception {
        assertEquals("", invokeGetStr(new JsonObject(), "missing"));
    }

    @Test
    void getStr_nullJsonValue_returnsEmpty() throws Exception {
        JsonObject obj = new JsonObject();
        obj.add("key", JsonNull.INSTANCE);
        assertEquals("", invokeGetStr(obj, "key"));
    }

    @Test
    void getStr_presentValue_returnsString() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("name", "Budget.xlsx");
        assertEquals("Budget.xlsx", invokeGetStr(obj, "name"));
    }

    // ── getObj ───────────────────────────────────────────────────────────────

    @Test
    void getObj_nullObject_returnsEmptyJsonObject() throws Exception {
        assertTrue(invokeGetObj(null, "params").entrySet().isEmpty());
    }

    @Test
    void getObj_missingKey_returnsEmptyJsonObject() throws Exception {
        assertTrue(invokeGetObj(new JsonObject(), "params").entrySet().isEmpty());
    }

    @Test
    void getObj_nonObjectValue_returnsEmptyJsonObject() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("key", "string");
        assertTrue(invokeGetObj(obj, "key").entrySet().isEmpty());
    }

    @Test
    void getObj_presentObject_returnsIt() throws Exception {
        JsonObject inner = new JsonObject();
        inner.addProperty("foo", "bar");
        JsonObject obj = new JsonObject();
        obj.add("params", inner);
        assertEquals("bar", invokeGetObj(obj, "params").get("foo").getAsString());
    }

    // ── buildSearchBody ──────────────────────────────────────────────────────

    @Test
    void buildSearchBody_queryStringIsSet() throws Exception {
        JsonObject parsed = GSON.fromJson(invokeBuildSearchBody("FileType:xlsx", 10), JsonObject.class);
        String qs = parsed.getAsJsonArray("requests").get(0)
                .getAsJsonObject().getAsJsonObject("query").get("queryString").getAsString();
        assertEquals("FileType:xlsx", qs);
    }

    @Test
    void buildSearchBody_entityTypeIsDriveItem() throws Exception {
        JsonObject parsed = GSON.fromJson(invokeBuildSearchBody("test", 25), JsonObject.class);
        assertEquals("driveItem", parsed.getAsJsonArray("requests").get(0)
                .getAsJsonObject().getAsJsonArray("entityTypes").get(0).getAsString());
    }

    @Test
    void buildSearchBody_sizeIsCorrect() throws Exception {
        JsonObject parsed = GSON.fromJson(invokeBuildSearchBody("test", 50), JsonObject.class);
        assertEquals(50, parsed.getAsJsonArray("requests").get(0)
                .getAsJsonObject().get("size").getAsInt());
    }

    @Test
    void buildSearchBody_fromIsZero() throws Exception {
        JsonObject parsed = GSON.fromJson(invokeBuildSearchBody("test", 50), JsonObject.class);
        assertEquals(0, parsed.getAsJsonArray("requests").get(0)
                .getAsJsonObject().get("from").getAsInt());
    }

    // ── strProp / buildTool / buildSchema ─────────────────────────────────────

    @Test
    void strProp_typeIsString() throws Exception {
        assertEquals("string", invokeStrProp("desc").get("type").getAsString());
    }

    @Test
    void strProp_descriptionIsSet() throws Exception {
        assertEquals("my description", invokeStrProp("my description").get("description").getAsString());
    }

    @Test
    void buildTool_nameIsSet() throws Exception {
        assertEquals("my_tool", invokeBuildTool("my_tool", "desc", new JsonObject())
                .get("name").getAsString());
    }

    @Test
    void buildTool_descriptionIsSet() throws Exception {
        assertEquals("desc", invokeBuildTool("tool", "desc", new JsonObject())
                .get("description").getAsString());
    }

    @Test
    void buildTool_inputSchemaIsSet() throws Exception {
        JsonObject schema = new JsonObject();
        schema.addProperty("type", "object");
        assertNotNull(invokeBuildTool("tool", "desc", schema).get("inputSchema"));
    }

    @Test
    void buildSchema_withRequired_containsRequiredArray() throws Exception {
        JsonObject schema = invokeBuildSchema(List.of("fileId"), Map.of());
        assertTrue(schema.has("required"));
        assertEquals("fileId", schema.getAsJsonArray("required").get(0).getAsString());
    }

    @Test
    void buildSchema_emptyRequired_noRequiredKey() throws Exception {
        assertFalse(invokeBuildSchema(List.of(), Map.of()).has("required"));
    }

    @Test
    void buildSchema_typeIsObject() throws Exception {
        assertEquals("object", invokeBuildSchema(List.of(), Map.of()).get("type").getAsString());
    }

    // ── mcpError / toolError / toolResult ────────────────────────────────────

    @Test
    void mcpError_codeAndMessageAreSet() throws Exception {
        JsonObject json = parseEntity(invokeMcpError(new JsonPrimitive(1), -32601, "Not found"));
        assertEquals(-32601, json.getAsJsonObject("error").get("code").getAsInt());
        assertEquals("Not found", json.getAsJsonObject("error").get("message").getAsString());
    }

    @Test
    void mcpError_nullId_noIdInResponse() throws Exception {
        JsonObject json = parseEntity(invokeMcpError(null, -32700, "Parse error"));
        assertFalse(json.has("id"));
    }

    @Test
    void mcpError_jsonrpcIs2dot0() throws Exception {
        JsonObject json = parseEntity(invokeMcpError(new JsonPrimitive(1), -32603, "err"));
        assertEquals("2.0", json.get("jsonrpc").getAsString());
    }

    @Test
    void toolError_isErrorTrue() throws Exception {
        JsonObject result = toolErrorResult("Something went wrong");
        assertTrue(result.get("isError").getAsBoolean());
    }

    @Test
    void toolError_contentTextStartsWithError() throws Exception {
        assertTrue(contentText(toolErrorResult("Bad input")).startsWith("Error:"));
    }

    @Test
    void toolError_contentTypeIsText() throws Exception {
        JsonObject result = toolErrorResult("msg");
        assertEquals("text", result.getAsJsonArray("content").get(0)
                .getAsJsonObject().get("type").getAsString());
    }

    @Test
    void toolResult_contentTypeIsText() throws Exception {
        JsonObject result = toolResultResult("Hello");
        assertEquals("text", result.getAsJsonArray("content").get(0)
                .getAsJsonObject().get("type").getAsString());
    }

    @Test
    void toolResult_contentTextIsSet() throws Exception {
        assertEquals("Hello", contentText(toolResultResult("Hello")));
    }

    @Test
    void toolResult_noIsErrorField() throws Exception {
        assertFalse(toolResultResult("ok").has("isError"));
    }

    // ── resolvedPath ─────────────────────────────────────────────────────────

    @Test
    void resolvedPath_nullAttrs_returnsPathUnchanged() throws Exception {
        String result = invokeResolvedPath(resource, "/me/drive/root/children");
        assertEquals("/me/drive/root/children", result);
    }

    @Test
    void resolvedPath_nonPrivateAuth_returnsPathUnchanged() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        lenient().when(attrs.getAuthType()).thenReturn("PUBLIC");
        lenient().when(attributeStore.load("test-invoker")).thenReturn(attrs);
        assertEquals("/me/drive/root/children",
                invokeResolvedPath(resource, "/me/drive/root/children"));
    }

    @Test
    void resolvedPath_privateAuthWithEmail_replacesMeWithUsers() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PRIVATE");
        when(attrs.getEmail()).thenReturn("user@example.com");
        when(attributeStore.load("test-invoker")).thenReturn(attrs);
        assertEquals("/users/user@example.com/drive/root/children",
                invokeResolvedPath(resource, "/me/drive/root/children"));
    }

    @Test
    void resolvedPath_privateAuthBlankEmail_returnsPathUnchanged() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PRIVATE");
        when(attrs.getEmail()).thenReturn("   ");
        when(attributeStore.load("test-invoker")).thenReturn(attrs);
        assertEquals("/me/drive/root/children",
                invokeResolvedPath(resource, "/me/drive/root/children"));
    }

    @Test
    void resolvedPath_privateAuthNullEmail_returnsPathUnchanged() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PRIVATE");
        when(attrs.getEmail()).thenReturn(null);
        when(attributeStore.load("test-invoker")).thenReturn(attrs);
        assertEquals("/me/drive/root/children",
                invokeResolvedPath(resource, "/me/drive/root/children"));
    }

    @Test
    void resolvedPath_privateAuth_trailingSlash_replacedCorrectly() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PRIVATE");
        when(attrs.getEmail()).thenReturn("admin@contoso.com");
        when(attributeStore.load("test-invoker")).thenReturn(attrs);
        assertEquals("/users/admin@contoso.com/drive/items/xyz",
                invokeResolvedPath(resource, "/me/drive/items/xyz"));
    }

    // ── adaptTokenStore ───────────────────────────────────────────────────────

    @Test
    void adaptTokenStore_save_delegatesToRefreshTokenStorePut() throws Exception {
        ITokenStore store = invokeAdaptTokenStore(refreshTokenStore);
        store.save("k1", "v1");
        verify(refreshTokenStore).put("k1", "v1");
    }

    @Test
    void adaptTokenStore_get_delegatesToRefreshTokenStoreGet() throws Exception {
        when(refreshTokenStore.get("k1")).thenReturn("v1");
        ITokenStore store = invokeAdaptTokenStore(refreshTokenStore);
        assertEquals("v1", store.get("k1"));
    }

    @Test
    void adaptTokenStore_delete_delegatesToRefreshTokenStoreRemove() throws Exception {
        ITokenStore store = invokeAdaptTokenStore(refreshTokenStore);
        store.delete("k1");
        verify(refreshTokenStore).remove("k1");
    }

    @Test
    void adaptTokenStore_save_ioException_throwsUncheckedIOException() throws Exception {
        doThrow(new IOException("disk full")).when(refreshTokenStore).put(anyString(), anyString());
        ITokenStore store = invokeAdaptTokenStore(refreshTokenStore);
        assertThrows(UncheckedIOException.class, () -> store.save("k", "v"));
    }

    @Test
    void adaptTokenStore_get_ioException_throwsUncheckedIOException() throws Exception {
        when(refreshTokenStore.get(anyString())).thenThrow(new IOException("network error"));
        ITokenStore store = invokeAdaptTokenStore(refreshTokenStore);
        assertThrows(UncheckedIOException.class, () -> store.get("k"));
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private JsonObject parseEntity(Response resp) {
        return GSON.fromJson((String) resp.getEntity(), JsonObject.class);
    }

    /** Calls handle() with a minimal JSON-RPC request for the given method and returns result. */
    private JsonObject invokeHandle(String method) {
        String req = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"" + method + "\"}";
        return parseEntity(resource.handle(req)).getAsJsonObject("result");
    }

    /** Calls handle() with tools/call and returns the result object. */
    private JsonObject toolsCallResult(String toolName, String arguments) {
        String req = "{\"jsonrpc\":\"2.0\",\"id\":6,\"method\":\"tools/call\","
                + "\"params\":{\"name\":\"" + toolName + "\",\"arguments\":" + arguments + "}}";
        return parseEntity(resource.handle(req)).getAsJsonObject("result");
    }

    private String contentText(JsonObject result) {
        return result.getAsJsonArray("content").get(0).getAsJsonObject().get("text").getAsString();
    }

    private List<String> toolNames(JsonObject result) {
        List<String> names = new java.util.ArrayList<>();
        result.getAsJsonArray("tools").forEach(t ->
                names.add(t.getAsJsonObject().get("name").getAsString()));
        return names;
    }

    private String searchJson(String name, String hitId) {
        return "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":\"" + hitId + "\","
                + "\"resource\":{"
                + "  \"name\":\"" + name + "\","
                + "  \"lastModifiedDateTime\":\"2024-05-15T10:00:00Z\","
                + "  \"size\":\"2048\","
                + "  \"webUrl\":\"https://onedrive.com/1\""
                + "}}]}]}]}";
    }

    private String singleItemJson() {
        return "{\"name\":\"Budget.xlsx\",\"id\":\"abc-123\","
                + "\"size\":\"1048576\",\"lastModifiedDateTime\":\"2024-05-15T10:00:00Z\","
                + "\"webUrl\":\"https://onedrive.com/file\"}";
    }

    private static JsonObject buildIdentitySet(String fieldName, String displayName) {
        JsonObject user = new JsonObject();
        user.addProperty("displayName", displayName);
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", user);
        JsonObject parent = new JsonObject();
        parent.add(fieldName, identitySet);
        return parent;
    }

    // ── Reflection invokers ───────────────────────────────────────────────────

    private static String invokeFormatSearchResponse(String json, String label) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "formatSearchResponse", String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, json, label);
    }

    private static String invokeFormatSingleItem(String json) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod("formatSingleItem", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, json);
    }

    private static String invokeFormatWorksheetsResponse(String json) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod("formatWorksheetsResponse", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, json);
    }

    private static String invokeFormatWorksheetData(String json, String sheetName) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "formatWorksheetData", String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, json, sheetName);
    }

    private static void invokeAppendRow(StringBuilder sb, List<String> row, int[] widths) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "appendRow", StringBuilder.class, java.util.List.class, int[].class);
        m.setAccessible(true);
        m.invoke(null, sb, row, widths);
    }

    private static void invokeAppendSeparator(StringBuilder sb, int[] widths) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "appendSeparator", StringBuilder.class, int[].class);
        m.setAccessible(true);
        m.invoke(null, sb, widths);
    }

    private static String invokePad(String s, int width) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod("pad", String.class, int.class);
        m.setAccessible(true);
        return (String) m.invoke(null, s, width);
    }

    private static String invokeWrapGraphError(int status, String body) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "wrapGraphError", int.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, status, body);
    }

    private static String invokeExtractDisplayName(JsonObject parent, String fieldName) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "extractDisplayName", JsonObject.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, parent, fieldName);
    }

    private static String invokeTruncateDate(String dateTime) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod("truncateDate", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, dateTime);
    }

    private static String invokeFormatSize(String sizeStr) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod("formatSize", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, sizeStr);
    }

    private static String invokeGetStr(JsonObject obj, String key) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "getStr", JsonObject.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, obj, key);
    }

    private static JsonObject invokeGetObj(JsonObject obj, String key) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "getObj", JsonObject.class, String.class);
        m.setAccessible(true);
        return (JsonObject) m.invoke(null, obj, key);
    }

    private static String invokeBuildSearchBody(String queryString, int size) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "buildSearchBody", String.class, int.class);
        m.setAccessible(true);
        return (String) m.invoke(null, queryString, size);
    }

    private static JsonObject invokeStrProp(String description) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod("strProp", String.class);
        m.setAccessible(true);
        return (JsonObject) m.invoke(null, description);
    }

    private static JsonObject invokeBuildTool(String name, String description, JsonObject schema) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "buildTool", String.class, String.class, JsonObject.class);
        m.setAccessible(true);
        return (JsonObject) m.invoke(null, name, description, schema);
    }

    private static JsonObject invokeBuildSchema(List<String> required, Map<String, JsonObject> props) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "buildSchema", List.class, Map.class);
        m.setAccessible(true);
        return (JsonObject) m.invoke(null, required, props);
    }

    private static Response invokeMcpError(com.google.gson.JsonElement id, int code, String message) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "mcpError", com.google.gson.JsonElement.class, int.class, String.class);
        m.setAccessible(true);
        return (Response) m.invoke(null, id, code, message);
    }

    private static JsonObject toolErrorResult(String message) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "toolError", com.google.gson.JsonElement.class, String.class);
        m.setAccessible(true);
        Response resp = (Response) m.invoke(null, new JsonPrimitive(1), message);
        return GSON.fromJson((String) resp.getEntity(), JsonObject.class).getAsJsonObject("result");
    }

    private static JsonObject toolResultResult(String text) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "toolResult", com.google.gson.JsonElement.class, String.class);
        m.setAccessible(true);
        Response resp = (Response) m.invoke(null, new JsonPrimitive(1), text);
        return GSON.fromJson((String) resp.getEntity(), JsonObject.class).getAsJsonObject("result");
    }

    private static String invokeResolvedPath(ExcelMcpProtocolResource resource, String path) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod("resolvedPath", String.class);
        m.setAccessible(true);
        return (String) m.invoke(resource, path);
    }

    private static ITokenStore invokeAdaptTokenStore(RefreshTokenStore store) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "adaptTokenStore", RefreshTokenStore.class);
        m.setAccessible(true);
        return (ITokenStore) m.invoke(null, store);
    }
}
