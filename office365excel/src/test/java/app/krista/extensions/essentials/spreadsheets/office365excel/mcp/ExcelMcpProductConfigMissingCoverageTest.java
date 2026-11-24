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
import com.google.gson.JsonArray;
import com.google.gson.JsonNull;
import com.google.gson.JsonObject;
import com.google.gson.JsonPrimitive;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Supplemental tests for {@link ExcelMcpProductConfig} to reach 100% coverage
 * on all branches not exercised by {@link ExcelMcpProductConfigTest}.
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>{@code parseAllFieldsResponse} — all branches</li>
 *   <li>{@code parseSearchResponse} — missing branches
 *       (no-value, no-hitsContainers, no-total, no-hits, no-resource, empty hitId)</li>
 *   <li>{@code parseStandardResponse} — missing branches</li>
 *   <li>{@code buildSearchRequestBody} — allFields=true (ignored but branch covered)</li>
 *   <li>{@code toDisplayValue} — null, JsonNull, DateTime, date, size, object fields,
 *       unknown object, JsonArray</li>
 *   <li>{@code toFreeFormKey} — all switch cases and default</li>
 *   <li>{@code extractUserDisplayName} — null, missing user, null user,
 *       empty displayName, present displayName</li>
 *   <li>{@code getString} — null obj, missing key, null value, present value</li>
 *   <li>{@code formatSize} — null, empty, bytes, KB, MB, non-numeric</li>
 * </ul>
 */
class ExcelMcpProductConfigMissingCoverageTest {

    private ExcelMcpProductConfig config;

    @BeforeEach
    void setUp() {
        config = new ExcelMcpProductConfig();
    }

    // ── parseAllFieldsResponse ────────────────────────────────────────────────

    @Test
    void parseAllFieldsResponse_emptyValue_returnsEmptyList() {
        List<Map<String, String>> rows = config.parseAllFieldsResponse("{\"value\":[]}", "");
        assertTrue(rows.isEmpty());
    }

    @Test
    void parseAllFieldsResponse_noValueKey_returnsEmptyList() {
        List<Map<String, String>> rows = config.parseAllFieldsResponse("{}", "");
        assertTrue(rows.isEmpty());
    }

    @Test
    void parseAllFieldsResponse_singleHit_containsName() {
        List<Map<String, String>> rows = config.parseAllFieldsResponse(hitJson("Budget.xlsx", ""), "");
        assertEquals(1, rows.size());
        assertEquals("Budget.xlsx", rows.get(0).get("Name"));
    }

    @Test
    void parseAllFieldsResponse_hitId_storedAsFileId() {
        List<Map<String, String>> rows = config.parseAllFieldsResponse(hitJson("File.xlsx", "hit-001"), "");
        assertEquals("hit-001", rows.get(0).get("fileId"));
    }

    @Test
    void parseAllFieldsResponse_emptyHitId_notAddedAsFileId() {
        // hitId absent → getString returns "" → not put into row
        List<Map<String, String>> rows = config.parseAllFieldsResponse(hitJson("File.xlsx", ""), "");
        assertFalse(rows.get(0).containsKey("fileId"),
                "Empty hitId must not be stored as fileId");
    }

    @Test
    void parseAllFieldsResponse_keyAndTitleAlwaysPresent() {
        // toFreeFormKey("name") → "Name" (PascalCase), so row has "Name" not "name".
        // getOrDefault("name", "") always returns "" — key/title are always empty strings
        // but the fields are always added to the row.
        List<Map<String, String>> rows = config.parseAllFieldsResponse(hitJson("Report.xlsx", ""), "");
        assertTrue(rows.get(0).containsKey("key"),   "'key' must always be present in the row map");
        assertTrue(rows.get(0).containsKey("title"), "'title' must always be present in the row map");
        // The underlying name is stored as "Name" (PascalCase), so lowercase "name"
        // lookup via getOrDefault returns the default ""
        assertEquals("", rows.get(0).get("key"));
        assertEquals("", rows.get(0).get("title"));
    }

    @Test
    void parseAllFieldsResponse_keyAndTitleEmptyWhenNoName() {
        // resource has no "name" field → getOrDefault("name","") returns "" in both cases
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{\"webUrl\":\"https://example.com\"}"
                + "}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals("", rows.get(0).get("key"));
        assertEquals("", rows.get(0).get("title"));
    }

    @Test
    void parseAllFieldsResponse_summaryPresentInHit_storedAsSummary() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"summary\":\"This is a summary\","
                + "\"resource\":{\"name\":\"File.xlsx\"}"
                + "}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals("This is a summary", rows.get(0).get("summary"));
    }

    @Test
    void parseAllFieldsResponse_summaryAbsent_notStoredAsSummary() {
        List<Map<String, String>> rows = config.parseAllFieldsResponse(hitJson("File.xlsx", ""), "");
        assertFalse(rows.get(0).containsKey("summary"),
                "Absent summary must not produce a summary key");
    }

    @Test
    void parseAllFieldsResponse_emptySummary_notStoredAsSummary() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"summary\":\"\","
                + "\"resource\":{\"name\":\"File.xlsx\"}"
                + "}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertFalse(rows.get(0).containsKey("summary"),
                "Empty summary must not produce a summary key");
    }

    @Test
    void parseAllFieldsResponse_multipleHits_allReturned() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":["
                + "{\"resource\":{\"name\":\"A.xlsx\"}},"
                + "{\"resource\":{\"name\":\"B.xlsx\"}}"
                + "]}]}]}";
        assertEquals(2, config.parseAllFieldsResponse(json, "").size());
    }

    @Test
    void parseAllFieldsResponse_noHitsContainersKey_returnsEmptyList() {
        String json = "{\"value\":[{}]}";
        assertTrue(config.parseAllFieldsResponse(json, "").isEmpty());
    }

    @Test
    void parseAllFieldsResponse_noHitsKey_returnsEmptyList() {
        String json = "{\"value\":[{\"hitsContainers\":[{}]}]}";
        assertTrue(config.parseAllFieldsResponse(json, "").isEmpty());
    }

    @Test
    void parseAllFieldsResponse_noResourceKey_returnsRowWithKeyAndTitle() {
        // hit has no "resource" → empty JsonObject used, no fields added, only key/title
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals(1, rows.size());
        assertEquals("", rows.get(0).get("key"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithLastModifiedDateTime_truncatesDate() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"File.xlsx\","
                + "  \"lastModifiedDateTime\":\"2024-07-01T12:00:00Z\""
                + "}}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals("2024-07-01", rows.get(0).get("LastModified"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithSizeBytes_formatsAsKB() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{\"name\":\"F.xlsx\",\"size\":\"4096\"}"
                + "}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals("4 KB", rows.get(0).get("Size"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithLastModifiedBy_extractsDisplayName() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"lastModifiedBy\":{\"user\":{\"displayName\":\"Carol\"}}"
                + "}}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals("Carol", rows.get(0).get("ModifiedBy"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithCreatedBy_extractsDisplayName() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"createdBy\":{\"user\":{\"displayName\":\"Dave\"}}"
                + "}}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals("Dave", rows.get(0).get("CreatedBy"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithParentReference_extractsFolderName() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"parentReference\":{\"name\":\"Documents\"}"
                + "}}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertEquals("Documents", rows.get(0).get("Folder"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithFile_extractsMimeType() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"file\":{\"mimeType\":\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\"}"
                + "}}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertTrue(rows.get(0).get("MimeType").contains("spreadsheetml"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithUnknownObjectField_skipped() {
        // Unknown object field → toDisplayValue returns null → not added
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"thumbnails\":{\"large\":{\"url\":\"https://thumb\"}}"
                + "}}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertFalse(rows.get(0).containsKey("thumbnails"));
    }

    @Test
    void parseAllFieldsResponse_resourceWithArrayField_skipped() {
        // JsonArray → toDisplayValue returns null → not added
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"tags\":[\"finance\",\"2024\"]"
                + "}}]}]}]}";
        List<Map<String, String>> rows = config.parseAllFieldsResponse(json, "");
        assertFalse(rows.get(0).containsKey("tags"));
    }

    // ── parseSearchResponse — missing branches ────────────────────────────────

    @Test
    void parseSearchResponse_noValueKey_returnsZeroRows() {
        McpSearchResult result = config.parseSearchResponse("{}", "");
        assertEquals(0, result.shown());
        assertEquals(0, result.total());
    }

    @Test
    void parseSearchResponse_noHitsContainersKey_returnsZeroRows() {
        McpSearchResult result = config.parseSearchResponse("{\"value\":[{}]}", "");
        assertEquals(0, result.shown());
    }

    @Test
    void parseSearchResponse_noTotalKey_totalIsZero() {
        // hitsContainer without "total" → total += 0
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{\"name\":\"F.xlsx\"}"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals(0, result.total());
        assertEquals(1, result.shown());
    }

    @Test
    void parseSearchResponse_noHitsKey_returnsZeroRows() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":0}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals(0, result.shown());
    }

    @Test
    void parseSearchResponse_noResourceKey_rowAddedWithHitIdOnly() {
        // hit has no "resource" → empty JsonObject, no resource fields; hitId still stored
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"hitId\":\"only-id\""
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals(1, result.shown());
        assertEquals("only-id", result.rows().get(0).get("fileId"));
    }

    @Test
    void parseSearchResponse_emptyHitId_notStoredAsFileId() {
        // hitId is empty string → getString returns "" → not added
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"hitId\":\"\","
                + "\"resource\":{\"name\":\"F.xlsx\"}"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertFalse(result.rows().get(0).containsKey("fileId"));
    }

    @Test
    void parseSearchResponse_lastModifiedBy_storedAsModifiedBy() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"lastModifiedBy\":{\"user\":{\"displayName\":\"Eve\"}}"
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("Eve", result.rows().get(0).get("ModifiedBy"));
    }

    @Test
    void parseSearchResponse_createdBy_storedAsCreatedBy() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"createdBy\":{\"user\":{\"displayName\":\"Frank\"}}"
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("Frank", result.rows().get(0).get("CreatedBy"));
    }

    @Test
    void parseSearchResponse_createdDateTime_storedAsCreated() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"createdDateTime\":\"2024-01-10T08:00:00Z\""
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("2024-01-10", result.rows().get(0).get("Created"));
    }

    @Test
    void parseSearchResponse_parentReference_storedAsFolder() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"parentReference\":{\"name\":\"Finance\"}"
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("Finance", result.rows().get(0).get("Folder"));
    }

    @Test
    void parseSearchResponse_fileField_storedAsMimeType() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"file\":{\"mimeType\":\"application/vnd.ms-excel\"}"
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("application/vnd.ms-excel", result.rows().get(0).get("MimeType"));
    }

    @Test
    void parseSearchResponse_idField_storedAsFileId() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"id\":\"drive-id-999\""
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("drive-id-999", result.rows().get(0).get("FileId"));
    }

    @Test
    void parseSearchResponse_unknownObjectField_skipped() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"unknownObj\":{\"a\":\"b\"}"
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertFalse(result.rows().get(0).containsKey("unknownObj"));
    }

    @Test
    void parseSearchResponse_arrayField_skipped() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"tags\":[\"a\",\"b\"]"
                + "}}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertFalse(result.rows().get(0).containsKey("tags"));
    }

    @Test
    void parseSearchResponse_sizeFormatsMB() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{\"name\":\"F.xlsx\",\"size\":\"2097152\"}"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("2.0 MB", result.rows().get(0).get("Size"));
    }

    @Test
    void parseSearchResponse_sizeFormatsBytes() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"resource\":{\"name\":\"F.xlsx\",\"size\":\"512\"}"
                + "}]}]}]}";
        McpSearchResult result = config.parseSearchResponse(json, "");
        assertEquals("512 B", result.rows().get(0).get("Size"));
    }

    // ── parseStandardResponse — missing branches ──────────────────────────────

    @Test
    void parseStandardResponse_noValueKey_returnsEmptyList() {
        assertTrue(config.parseStandardResponse("{}", "").isEmpty());
    }

    @Test
    void parseStandardResponse_noHitsContainersKey_returnsEmptyList() {
        assertTrue(config.parseStandardResponse("{\"value\":[{}]}", "").isEmpty());
    }

    @Test
    void parseStandardResponse_noHitsKey_returnsEmptyList() {
        String json = "{\"value\":[{\"hitsContainers\":[{}]}]}";
        assertTrue(config.parseStandardResponse(json, "").isEmpty());
    }

    @Test
    void parseStandardResponse_noResourceKey_returnsEmptyFreeForm() {
        // hit has no "resource" → empty JsonObject → no fields → FreeForm empty, hitId empty
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertEquals(1, rows.size()); // row still added, just empty
    }

    @Test
    void parseStandardResponse_emptyHitId_notAddedToFreeForm() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":\"\","
                + "\"resource\":{\"name\":\"F.xlsx\"}"
                + "}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertNull(rows.get(0).get("FileId"),
                "Empty hitId must not produce a FileId FreeForm entry");
    }

    @Test
    void parseStandardResponse_lastModifiedBy_storedAsModifiedBy() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"lastModifiedBy\":{\"user\":{\"displayName\":\"Grace\"}}"
                + "}}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertNotNull(rows.get(0).get("ModifiedBy"));
    }

    @Test
    void parseStandardResponse_createdBy_storedAsCreatedBy() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"createdBy\":{\"user\":{\"displayName\":\"Heidi\"}}"
                + "}}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertNotNull(rows.get(0).get("CreatedBy"));
    }

    @Test
    void parseStandardResponse_parentReference_storedAsFolder() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"parentReference\":{\"name\":\"HR\"}"
                + "}}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertNotNull(rows.get(0).get("Folder"));
    }

    @Test
    void parseStandardResponse_fileField_storedAsMimeType() {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"file\":{\"mimeType\":\"application/vnd.ms-excel\"}"
                + "}}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertNotNull(rows.get(0).get("MimeType"));
    }

    @Test
    void parseStandardResponse_emptyDisplayValueField_notAddedToFreeForm() {
        // null val → toDisplayValue returns null → not put in FreeForm
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"resource\":{"
                + "  \"name\":\"F.xlsx\","
                + "  \"thumbnails\":{}"
                + "}}]}]}]}";
        List<FreeForm> rows = config.parseStandardResponse(json, "");
        assertNull(rows.get(0).get("thumbnails"));
    }

    // ── buildSearchRequestBody — allFields branch ─────────────────────────────

    @Test
    void buildSearchRequestBody_allFieldsTrue_stillProducesValidBody() {
        // allFields is unused in the implementation; just ensures branch is reached
        String body = config.buildSearchRequestBody("budget", 10, true);
        assertNotNull(body);
        assertTrue(body.contains("driveItem"));
    }

    @Test
    void buildSearchRequestBody_queryAlreadyContainsFileTypeCaseInsensitive() {
        // "filetype:" (lowercase) → toLowerCase contains → no append
        String body = config.buildSearchRequestBody("filetype:xlsx custom", 10, false);
        int count = body.split("(?i)FileType:", -1).length - 1;
        assertEquals(1, count, "FileType filter must not be appended when already present (case-insensitive)");
    }

    // ── toDisplayValue — via reflection ───────────────────────────────────────

    @Test
    void toDisplayValue_nullVal_returnsNull() throws Exception {
        assertNull(invokeToDisplayValue("name", null, new JsonObject()));
    }

    @Test
    void toDisplayValue_jsonNull_returnsNull() throws Exception {
        assertNull(invokeToDisplayValue("name", JsonNull.INSTANCE, new JsonObject()));
    }

    @Test
    void toDisplayValue_primitiveDateTimeKeyWithT_truncatesToDate() throws Exception {
        assertEquals("2024-05-15",
                invokeToDisplayValue("lastModifiedDateTime",
                        new JsonPrimitive("2024-05-15T10:00:00Z"), new JsonObject()));
    }

    @Test
    void toDisplayValue_primitiveDateKeyWithT_truncatesToDate() throws Exception {
        // key contains "date" (lowercase)
        assertEquals("2024-05-15",
                invokeToDisplayValue("somedate",
                        new JsonPrimitive("2024-05-15T10:00:00Z"), new JsonObject()));
    }

    @Test
    void toDisplayValue_primitiveDateTimeKeyNoT_returnedAsIs() throws Exception {
        assertEquals("2024-05-15",
                invokeToDisplayValue("lastModifiedDateTime",
                        new JsonPrimitive("2024-05-15"), new JsonObject()));
    }

    @Test
    void toDisplayValue_primitiveSizeKey_formatsSize() throws Exception {
        assertEquals("1 KB",
                invokeToDisplayValue("size", new JsonPrimitive("1024"), new JsonObject()));
    }

    @Test
    void toDisplayValue_primitiveOtherKey_returnedAsIs() throws Exception {
        assertEquals("Budget.xlsx",
                invokeToDisplayValue("name", new JsonPrimitive("Budget.xlsx"), new JsonObject()));
    }

    @Test
    void toDisplayValue_objectLastModifiedBy_extractsDisplayName() throws Exception {
        JsonObject identitySet = buildIdentitySet("Alice");
        assertEquals("Alice",
                invokeToDisplayValue("lastModifiedBy", identitySet, new JsonObject()));
    }

    @Test
    void toDisplayValue_objectCreatedBy_extractsDisplayName() throws Exception {
        JsonObject identitySet = buildIdentitySet("Bob");
        assertEquals("Bob",
                invokeToDisplayValue("createdBy", identitySet, new JsonObject()));
    }

    @Test
    void toDisplayValue_objectParentReference_returnsFolderName() throws Exception {
        JsonObject parentRef = new JsonObject();
        parentRef.addProperty("name", "Finance");
        assertEquals("Finance",
                invokeToDisplayValue("parentReference", parentRef, new JsonObject()));
    }

    @Test
    void toDisplayValue_objectFile_returnsMimeType() throws Exception {
        JsonObject fileObj = new JsonObject();
        fileObj.addProperty("mimeType", "application/vnd.ms-excel");
        assertEquals("application/vnd.ms-excel",
                invokeToDisplayValue("file", fileObj, new JsonObject()));
    }

    @Test
    void toDisplayValue_unknownObjectKey_returnsNull() throws Exception {
        JsonObject unknownObj = new JsonObject();
        unknownObj.addProperty("foo", "bar");
        assertNull(invokeToDisplayValue("unknownKey", unknownObj, new JsonObject()),
                "Unknown object-type fields must return null (skipped)");
    }

    @Test
    void toDisplayValue_jsonArray_returnsNull() throws Exception {
        assertNull(invokeToDisplayValue("tags", new JsonArray(), new JsonObject()),
                "JsonArray values must return null (skipped)");
    }

    // ── toFreeFormKey — via reflection ────────────────────────────────────────

    @Test
    void toFreeFormKey_name_returnsName() throws Exception {
        assertEquals("Name", invokeToFreeFormKey("name"));
    }

    @Test
    void toFreeFormKey_webUrl_returnsLink() throws Exception {
        assertEquals("Link", invokeToFreeFormKey("webUrl"));
    }

    @Test
    void toFreeFormKey_lastModifiedDateTime_returnsLastModified() throws Exception {
        assertEquals("LastModified", invokeToFreeFormKey("lastModifiedDateTime"));
    }

    @Test
    void toFreeFormKey_createdDateTime_returnsCreated() throws Exception {
        assertEquals("Created", invokeToFreeFormKey("createdDateTime"));
    }

    @Test
    void toFreeFormKey_size_returnsSize() throws Exception {
        assertEquals("Size", invokeToFreeFormKey("size"));
    }

    @Test
    void toFreeFormKey_id_returnsFileId() throws Exception {
        assertEquals("FileId", invokeToFreeFormKey("id"));
    }

    @Test
    void toFreeFormKey_lastModifiedBy_returnsModifiedBy() throws Exception {
        assertEquals("ModifiedBy", invokeToFreeFormKey("lastModifiedBy"));
    }

    @Test
    void toFreeFormKey_createdBy_returnsCreatedBy() throws Exception {
        assertEquals("CreatedBy", invokeToFreeFormKey("createdBy"));
    }

    @Test
    void toFreeFormKey_parentReference_returnsFolder() throws Exception {
        assertEquals("Folder", invokeToFreeFormKey("parentReference"));
    }

    @Test
    void toFreeFormKey_file_returnsMimeType() throws Exception {
        assertEquals("MimeType", invokeToFreeFormKey("file"));
    }

    @Test
    void toFreeFormKey_unknownField_passesThrough() throws Exception {
        assertEquals("customField", invokeToFreeFormKey("customField"));
        assertEquals("worksheetId", invokeToFreeFormKey("worksheetId"));
    }

    // ── extractUserDisplayName — via reflection ───────────────────────────────

    @Test
    void extractUserDisplayName_null_returnsEmpty() throws Exception {
        assertEquals("", invokeExtractUserDisplayName(null));
    }

    @Test
    void extractUserDisplayName_noUserKey_returnsEmpty() throws Exception {
        assertEquals("", invokeExtractUserDisplayName(new JsonObject()));
    }

    @Test
    void extractUserDisplayName_nullUser_returnsEmpty() throws Exception {
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", JsonNull.INSTANCE);
        assertEquals("", invokeExtractUserDisplayName(identitySet));
    }

    @Test
    void extractUserDisplayName_userWithDisplayName_returnsName() throws Exception {
        assertEquals("Carol", invokeExtractUserDisplayName(buildIdentitySet("Carol")));
    }

    @Test
    void extractUserDisplayName_userWithEmptyDisplayName_returnsEmpty() throws Exception {
        JsonObject user = new JsonObject();
        user.addProperty("displayName", "");
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", user);
        assertEquals("", invokeExtractUserDisplayName(identitySet));
    }

    @Test
    void extractUserDisplayName_userWithNoDisplayNameKey_returnsEmpty() throws Exception {
        JsonObject user = new JsonObject();
        user.addProperty("email", "user@example.com"); // no displayName key
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", user);
        assertEquals("", invokeExtractUserDisplayName(identitySet));
    }

    // ── getString — via reflection ────────────────────────────────────────────

    @Test
    void getString_nullObject_returnsEmpty() throws Exception {
        assertEquals("", invokeGetString(null, "key"));
    }

    @Test
    void getString_missingKey_returnsEmpty() throws Exception {
        assertEquals("", invokeGetString(new JsonObject(), "missing"));
    }

    @Test
    void getString_nullJsonValue_returnsEmpty() throws Exception {
        JsonObject obj = new JsonObject();
        obj.add("key", JsonNull.INSTANCE);
        assertEquals("", invokeGetString(obj, "key"));
    }

    @Test
    void getString_presentValue_returnsString() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("name", "Budget.xlsx");
        assertEquals("Budget.xlsx", invokeGetString(obj, "name"));
    }

    // ── formatSize — via reflection ───────────────────────────────────────────

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
    void formatSize_exactly1024_returnsKB() throws Exception {
        assertEquals("1 KB", invokeFormatSize("1024"));
    }

    @Test
    void formatSize_between1KBand1MB_returnsKB() throws Exception {
        assertEquals("2 KB", invokeFormatSize("2048"));
    }

    @Test
    void formatSize_exactly1MB_returnsMB() throws Exception {
        assertEquals("1.0 MB", invokeFormatSize("1048576"));
    }

    @Test
    void formatSize_largeMB_returnsMB() throws Exception {
        String result = invokeFormatSize("10485760");
        assertTrue(result.endsWith("MB"));
    }

    @Test
    void formatSize_nonNumeric_returnsAsIs() throws Exception {
        assertEquals("unknown-size", invokeFormatSize("unknown-size"));
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /** Builds a one-hit search JSON with optional hitId. */
    private static String hitJson(String name, String hitId) {
        String hitIdField = hitId.isEmpty() ? "" : "\"hitId\":\"" + hitId + "\",";
        return "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + hitIdField
                + "\"resource\":{\"name\":\"" + name + "\"}"
                + "}]}]}]}";
    }

    private static JsonObject buildIdentitySet(String displayName) {
        JsonObject user = new JsonObject();
        user.addProperty("displayName", displayName);
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", user);
        return identitySet;
    }

    // ── Reflection invokers ───────────────────────────────────────────────────

    private static String invokeToDisplayValue(String key,
                                               com.google.gson.JsonElement val,
                                               JsonObject resource) throws Exception {
        Method m = ExcelMcpProductConfig.class.getDeclaredMethod(
                "toDisplayValue", String.class,
                com.google.gson.JsonElement.class, JsonObject.class);
        m.setAccessible(true);
        return (String) m.invoke(null, key, val, resource);
    }

    private static String invokeToFreeFormKey(String graphKey) throws Exception {
        Method m = ExcelMcpProductConfig.class.getDeclaredMethod("toFreeFormKey", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, graphKey);
    }

    private static String invokeExtractUserDisplayName(JsonObject identitySet) throws Exception {
        Method m = ExcelMcpProductConfig.class.getDeclaredMethod(
                "extractUserDisplayName", JsonObject.class);
        m.setAccessible(true);
        return (String) m.invoke(null, identitySet);
    }

    private static String invokeGetString(JsonObject obj, String key) throws Exception {
        Method m = ExcelMcpProductConfig.class.getDeclaredMethod(
                "getString", JsonObject.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, obj, key);
    }

    private static String invokeFormatSize(String sizeStr) throws Exception {
        Method m = ExcelMcpProductConfig.class.getDeclaredMethod("formatSize", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, sizeStr);
    }
}
