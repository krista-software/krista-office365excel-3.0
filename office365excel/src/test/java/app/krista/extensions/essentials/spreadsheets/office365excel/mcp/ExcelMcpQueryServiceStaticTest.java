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
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonNull;
import com.google.gson.JsonObject;
import com.google.gson.JsonPrimitive;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for the private static helper methods in {@link ExcelMcpQueryService}.
 *
 * <p>All private methods are accessed via reflection.</p>
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>{@code toFreeFormKey} — Graph field name → PascalCase FreeForm key</li>
 *   <li>{@code toDisplayString} — JSON value → display string (dates, sizes, objects)</li>
 *   <li>{@code flattenValueArray} — 2D cell value array → pipe/slash display string</li>
 *   <li>{@code extractFileName} — filename extraction from natural language question</li>
 *   <li>{@code parseGenericResponse} — single object, flat list, search, sharedWithMe</li>
 *   <li>{@code getStr} — safe null-handling JSON string extraction</li>
 *   <li>{@code blank} — null/empty/whitespace detection</li>
 * </ul>
 */
class ExcelMcpQueryServiceStaticTest {

    // ── toFreeFormKey ─────────────────────────────────────────────────────

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
    void toFreeFormKey_address_returnsCell() throws Exception {
        assertEquals("Cell", invokeToFreeFormKey("address"),
                "'address' (workbook range cell address) must map to 'Cell'");
    }

    @Test
    void toFreeFormKey_values_returnsValue() throws Exception {
        assertEquals("Value", invokeToFreeFormKey("values"),
                "'values' (workbook range data) must map to 'Value'");
    }

    @Test
    void toFreeFormKey_unknownField_passesThroughUnchanged() throws Exception {
        assertEquals("customField", invokeToFreeFormKey("customField"));
        assertEquals("worksheetId", invokeToFreeFormKey("worksheetId"));
    }

    // ── toDisplayString ───────────────────────────────────────────────────

    @Test
    void toDisplayString_jsonNull_returnsNull() throws Exception {
        assertNull(invokeToDisplayString("name", JsonNull.INSTANCE, new JsonObject()));
    }

    @Test
    void toDisplayString_simplePrimitive_returnsString() throws Exception {
        assertEquals("Budget.xlsx",
                invokeToDisplayString("name", new JsonPrimitive("Budget.xlsx"), new JsonObject()));
    }

    @Test
    void toDisplayString_lastModifiedDateTime_truncatesToDate() throws Exception {
        String result = invokeToDisplayString("lastModifiedDateTime",
                new JsonPrimitive("2024-05-15T10:30:00Z"), new JsonObject());
        assertEquals("2024-05-15", result,
                "DateTime values must be truncated to date portion only (no T or time)");
    }

    @Test
    void toDisplayString_createdDateTime_truncatesToDate() throws Exception {
        String result = invokeToDisplayString("createdDateTime",
                new JsonPrimitive("2024-01-01T00:00:00Z"), new JsonObject());
        assertEquals("2024-01-01", result);
    }

    @Test
    void toDisplayString_size_formatsBytes() throws Exception {
        assertEquals("512 B",
                invokeToDisplayString("size", new JsonPrimitive("512"), new JsonObject()));
    }

    @Test
    void toDisplayString_size_formatsKilobytes() throws Exception {
        assertEquals("2 KB",
                invokeToDisplayString("size", new JsonPrimitive("2048"), new JsonObject()));
    }

    @Test
    void toDisplayString_size_formatsMegabytes() throws Exception {
        assertEquals("2.0 MB",
                invokeToDisplayString("size", new JsonPrimitive("2097152"), new JsonObject()));
    }

    @Test
    void toDisplayString_lastModifiedBy_extractsDisplayName() throws Exception {
        JsonObject user = new JsonObject();
        user.addProperty("displayName", "Alice Smith");
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", user);
        assertEquals("Alice Smith",
                invokeToDisplayString("lastModifiedBy", identitySet, new JsonObject()));
    }

    @Test
    void toDisplayString_createdBy_extractsDisplayName() throws Exception {
        JsonObject user = new JsonObject();
        user.addProperty("displayName", "Bob Jones");
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", user);
        assertEquals("Bob Jones",
                invokeToDisplayString("createdBy", identitySet, new JsonObject()));
    }

    @Test
    void toDisplayString_parentReference_extractsName() throws Exception {
        JsonObject parentRef = new JsonObject();
        parentRef.addProperty("name", "Documents");
        assertEquals("Documents",
                invokeToDisplayString("parentReference", parentRef, new JsonObject()));
    }

    @Test
    void toDisplayString_file_extractsMimeType() throws Exception {
        JsonObject fileObj = new JsonObject();
        fileObj.addProperty("mimeType",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        assertEquals(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                invokeToDisplayString("file", fileObj, new JsonObject()));
    }

    @Test
    void toDisplayString_values_flattens2dArray() throws Exception {
        JsonArray row = new JsonArray();
        row.add("A1");
        row.add("B1");
        JsonArray values = new JsonArray();
        values.add(row);
        assertEquals("A1 | B1",
                invokeToDisplayString("values", values, new JsonObject()));
    }

    @Test
    void toDisplayString_unknownObjectKey_returnsNull() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("someField", "someValue");
        assertNull(invokeToDisplayString("unknownObjectKey", obj, new JsonObject()),
                "Unknown object-type fields must be skipped (return null)");
    }

    // ── flattenValueArray ─────────────────────────────────────────────────

    @Test
    void flattenValueArray_emptyArray_returnsNull() throws Exception {
        assertNull(invokeFlattenValueArray(new JsonArray()));
    }

    @Test
    void flattenValueArray_singleRow_joinsCellsWithPipe() throws Exception {
        JsonArray row = new JsonArray();
        row.add("Hello");
        row.add("World");
        JsonArray arr = new JsonArray();
        arr.add(row);
        assertEquals("Hello | World", invokeFlattenValueArray(arr));
    }

    @Test
    void flattenValueArray_multipleRows_joinsRowsWithSlash() throws Exception {
        JsonArray row1 = new JsonArray();
        row1.add("r1c1");
        row1.add("r1c2");
        JsonArray row2 = new JsonArray();
        row2.add("r2c1");
        row2.add("r2c2");
        JsonArray arr = new JsonArray();
        arr.add(row1);
        arr.add(row2);
        assertEquals("r1c1 | r1c2 / r2c1 | r2c2", invokeFlattenValueArray(arr));
    }

    @Test
    void flattenValueArray_singleCellRow_noSeparator() throws Exception {
        JsonArray row = new JsonArray();
        row.add("OnlyCell");
        JsonArray arr = new JsonArray();
        arr.add(row);
        assertEquals("OnlyCell", invokeFlattenValueArray(arr));
    }

    @Test
    void flattenValueArray_threeByThree_correctFormat() throws Exception {
        JsonArray arr = new JsonArray();
        for (int r = 1; r <= 2; r++) {
            JsonArray row = new JsonArray();
            for (int c = 1; c <= 3; c++) {
                row.add("r" + r + "c" + c);
            }
            arr.add(row);
        }
        String result = invokeFlattenValueArray(arr);
        assertTrue(result.contains("r1c1 | r1c2 | r1c3"),
                "First row must have all three cells joined by pipe");
        assertTrue(result.contains(" / "),
                "Multiple rows must be joined by slash separator");
    }

    // ── extractFileName ───────────────────────────────────────────────────

    @Test
    void extractFileName_xlsxFile_found() throws Exception {
        assertEquals("show me Budget.xlsx", invokeExtractFileName("show me Budget.xlsx please"));
    }

    @Test
    void extractFileName_xlsFile_found() throws Exception {
        assertEquals("open Data.xls", invokeExtractFileName("open Data.xls"));
    }

    @Test
    void extractFileName_xlsmFile_found() throws Exception {
        assertEquals("read Macro.xls", invokeExtractFileName("read Macro.xlsm"));
    }

    @Test
    void extractFileName_xlsbFile_found() throws Exception {
        assertEquals("get Binary.xls", invokeExtractFileName("get Binary.xlsb data"));
    }

    @Test
    void extractFileName_caseInsensitive_found() throws Exception {
        assertNotNull(invokeExtractFileName("open budget.XLSX"),
                "Filename extraction must be case-insensitive");
    }

    @Test
    void extractFileName_noExcelFile_returnsNull() throws Exception {
        assertNull(invokeExtractFileName("show me all my Excel files"),
                "Question without a specific filename must return null");
    }

    @Test
    void extractFileName_pdfFile_returnsNull() throws Exception {
        assertNull(invokeExtractFileName("open Report.pdf"),
                "Non-Excel file extensions must not match");
    }

    // ── parseGenericResponse ──────────────────────────────────────────────

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_singleObject_returnsOneRow() throws Exception {
        String json = "{\"name\":\"Budget.xlsx\",\"webUrl\":\"https://onedrive.com/1\"}";
        List<FreeForm> rows = (List<FreeForm>) invokeParseGenericResponse(json);
        assertEquals(1, rows.size(),
                "A JSON object with no 'value' key must be treated as a single-row result");
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_flatList_returnsAllRows() throws Exception {
        String json = "{\"value\":["
                + "{\"name\":\"File1.xlsx\",\"webUrl\":\"https://onedrive.com/1\"},"
                + "{\"name\":\"File2.xlsx\",\"webUrl\":\"https://onedrive.com/2\"}"
                + "]}";
        List<FreeForm> rows = (List<FreeForm>) invokeParseGenericResponse(json);
        assertEquals(2, rows.size());
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_emptyValueArray_returnsEmpty() throws Exception {
        String json = "{\"value\":[]}";
        List<FreeForm> rows = (List<FreeForm>) invokeParseGenericResponse(json);
        assertTrue(rows.isEmpty());
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_searchResponse_extractsFromHits() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"hitId\":\"abc-123\","
                + "\"resource\":{\"name\":\"Report.xlsx\",\"webUrl\":\"https://onedrive.com/r\"}"
                + "}]}]}]}";
        List<FreeForm> rows = (List<FreeForm>) invokeParseGenericResponse(json);
        assertEquals(1, rows.size());
        assertNotNull(rows.get(0).get("Name"),
                "Name must be extracted from resource inside hitsContainers");
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_searchResponse_hitIdStoredAsFileId() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"total\":1,\"hits\":[{"
                + "\"hitId\":\"file-id-xyz\","
                + "\"resource\":{\"name\":\"Report.xlsx\"}"
                + "}]}]}]}";
        List<FreeForm> rows = (List<FreeForm>) invokeParseGenericResponse(json);
        assertNotNull(rows.get(0).get("FileId"),
                "hitId must be stored as 'FileId' in the FreeForm row");
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_sharedWithMeItem_mergesRemoteItemFields() throws Exception {
        String json = "{\"value\":[{"
                + "\"id\":\"local-id\","
                + "\"remoteItem\":{"
                + "  \"name\":\"SharedFile.xlsx\","
                + "  \"webUrl\":\"https://sharepoint.com/file\","
                + "  \"size\":\"1024\""
                + "}"
                + "}]}";
        List<FreeForm> rows = (List<FreeForm>) invokeParseGenericResponse(json);
        assertEquals(1, rows.size());
        assertNotNull(rows.get(0).get("Name"),
                "Name from remoteItem must be merged into the FreeForm row for sharedWithMe items");
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_odataMetadataFieldsSkipped() throws Exception {
        String json = "{\"@odata.context\":\"https://graph.microsoft.com\","
                + "\"name\":\"Budget.xlsx\"}";
        List<FreeForm> rows = (List<FreeForm>) invokeParseGenericResponse(json);
        assertEquals(1, rows.size());
        assertNull(rows.get(0).get("@odata.context"),
                "@odata.* metadata fields must be skipped");
    }

    // ── getStr ─────────────────────────────────────────────────────────────

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

    // ── blank ─────────────────────────────────────────────────────────────

    @Test
    void blank_null_returnsTrue() throws Exception {
        assertTrue((boolean) invokeBlank(null));
    }

    @Test
    void blank_empty_returnsTrue() throws Exception {
        assertTrue((boolean) invokeBlank(""));
    }

    @Test
    void blank_whitespace_returnsTrue() throws Exception {
        assertTrue((boolean) invokeBlank("   "));
    }

    @Test
    void blank_nonBlank_returnsFalse() throws Exception {
        assertFalse((boolean) invokeBlank("Bearer token123"));
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private static String invokeToFreeFormKey(String graphKey) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("toFreeFormKey", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, graphKey);
    }

    private static String invokeToDisplayString(String key, JsonElement val,
                                                JsonObject parent) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "toDisplayString", String.class, JsonElement.class, JsonObject.class);
        m.setAccessible(true);
        return (String) m.invoke(null, key, val, parent);
    }

    private static String invokeFlattenValueArray(JsonArray arr) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "flattenValueArray", JsonArray.class);
        m.setAccessible(true);
        return (String) m.invoke(null, arr);
    }

    private static String invokeExtractFileName(String question) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("extractFileName", String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, question);
    }

    @SuppressWarnings("unchecked")
    private static List<FreeForm> invokeParseGenericResponse(String json) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "parseGenericResponse", String.class);
        m.setAccessible(true);
        return (List<FreeForm>) m.invoke(null, json);
    }

    private static String invokeGetStr(JsonObject obj, String key) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "getStr", JsonObject.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, obj, key);
    }

    private static boolean invokeBlank(String s) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("blank", String.class);
        m.setAccessible(true);
        return (boolean) m.invoke(null, s);
    }
}
