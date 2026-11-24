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

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.model.base.FreeForm;
import app.krista.sdk.msauth.extension.mcp.McpSessionKeys;
import app.krista.sdk.msauth.token.ITokenStore;
import com.google.gson.JsonArray;
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

import java.io.IOException;
import java.io.UncheckedIOException;
import java.lang.reflect.Method;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.*;

/**
 * Supplemental tests for {@link ExcelMcpQueryService} covering branches
 * not exercised by {@link ExcelMcpQueryServiceStaticTest} or
 * {@link ExcelMcpQueryServiceInstanceTest}.
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>{@code getProductConfig} and {@code getJaxrsId}</li>
 *   <li>{@code query()} — AI key set but not authenticated</li>
 *   <li>{@code resolvedPath} — PRIVATE auth with / without email</li>
 *   <li>{@code graphObjectToFreeForm} — RANGE_SKIP_FIELDS, odata skip, empty value omission</li>
 *   <li>{@code toDisplayString} — "date" key, non-numeric size, missing/null user,
 *       values with non-array row element</li>
 *   <li>{@code flattenValueArray} — null input, null cell value</li>
 *   <li>{@code parseGenericResponse} — flat list with non-object element,
 *       search with empty hitsContainers / empty hits</li>
 *   <li>{@code getJsonString} — all branches</li>
 *   <li>{@code adaptTokenStore} — save / get / delete and IOException wrapping</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class ExcelMcpQueryServiceMissingCoverageTest {

    @Mock private RefreshTokenStore         refreshTokenStore;
    @Mock private Invoker                   invoker;
    @Mock private OfficeExcelAttributeStore attributeStore;

    private ExcelMcpQueryService service;

    @BeforeEach
    void setUp() throws Exception {
        lenient().when(invoker.getInvokerId()).thenReturn("test-invoker");
        lenient().when(attributeStore.load("test-invoker")).thenReturn(null);
        lenient().when(refreshTokenStore.get(anyString())).thenReturn(null);
        service = new ExcelMcpQueryService(refreshTokenStore, invoker, attributeStore);
    }

    // ── getProductConfig / getJaxrsId ────────────────────────────────────────

    @Test
    void getProductConfig_returnsExcelMcpProductConfig() throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("getProductConfig");
        m.setAccessible(true);
        Object config = m.invoke(service);
        assertInstanceOf(ExcelMcpProductConfig.class, config);
    }

    @Test
    void getJaxrsId_returnsExcel() throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("getJaxrsId");
        m.setAccessible(true);
        assertEquals("excel", m.invoke(service));
    }

    // ── query() — AI key set but not authenticated ────────────────────────────

    @Test
    void query_aiKeySetNotAuthenticated_returnsFailure() throws Exception {
        // Return a non-blank claudeApiKey so AI key check passes, but attrs = null → not authenticated
        when(refreshTokenStore.get(McpSessionKeys.claudeApiKey("test-invoker"))).thenReturn("sk-test");
        ExtensionResponse resp = service.query("list my files");
        assertEquals(ExtensionResponse.Result.FAILURE, resp.getResult());
    }

    @Test
    void query_aiKeySetNotAuthenticated_errorContainsNotConnected() throws Exception {
        when(refreshTokenStore.get(McpSessionKeys.claudeApiKey("test-invoker"))).thenReturn("sk-test");
        ExtensionResponse resp = service.query("list my files");
        assertTrue(String.valueOf(resp.getError()).contains("Not connected"));
    }

    @Test
    void query_providerKeySetNotAuthenticated_returnsFailure() throws Exception {
        // Use provider + aiKey path: provider=openai, aiKey=sk-xxx
        when(refreshTokenStore.get(McpSessionKeys.aiProvider("test-invoker"))).thenReturn("openai");
        when(refreshTokenStore.get(McpSessionKeys.aiKey("test-invoker", "openai"))).thenReturn("sk-openai");
        ExtensionResponse resp = service.query("list files");
        assertEquals(ExtensionResponse.Result.FAILURE, resp.getResult());
    }

    // ── resolvedPath ─────────────────────────────────────────────────────────

    @Test
    void resolvedPath_privateAuthWithEmail_replacesMeWithUsers() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PRIVATE");
        when(attrs.getEmail()).thenReturn("user@contoso.com");
        when(attributeStore.load("test-invoker")).thenReturn(attrs);

        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "resolvedPath", String.class,
                app.krista.sdk.msauth.extension.mcp.DefaultMsAuthBridge.class);
        m.setAccessible(true);
        String result = (String) m.invoke(service, "/me/drive/root/children", null);
        assertEquals("/users/user@contoso.com/drive/root/children", result);
    }

    @Test
    void resolvedPath_privateAuthBlankEmail_returnsPathUnchanged() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PRIVATE");
        when(attrs.getEmail()).thenReturn("  ");
        when(attributeStore.load("test-invoker")).thenReturn(attrs);

        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "resolvedPath", String.class,
                app.krista.sdk.msauth.extension.mcp.DefaultMsAuthBridge.class);
        m.setAccessible(true);
        assertEquals("/me/drive/root/children",
                (String) m.invoke(service, "/me/drive/root/children", null));
    }

    @Test
    void resolvedPath_nonPrivateAuth_returnsPathUnchanged() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PUBLIC");
        when(attributeStore.load("test-invoker")).thenReturn(attrs);

        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "resolvedPath", String.class,
                app.krista.sdk.msauth.extension.mcp.DefaultMsAuthBridge.class);
        m.setAccessible(true);
        assertEquals("/me/drive/items/123",
                (String) m.invoke(service, "/me/drive/items/123", null));
    }

    // ── graphObjectToFreeForm ────────────────────────────────────────────────

    @Test
    void graphObjectToFreeForm_odataFieldSkipped() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("@odata.context", "https://graph.microsoft.com");
        obj.addProperty("name", "File.xlsx");
        FreeForm ff = invokeGraphObjectToFreeForm(obj);
        assertNull(ff.get("@odata.context"), "@odata fields must be skipped");
        assertNotNull(ff.get("Name"));
    }

    @Test
    void graphObjectToFreeForm_rangeSkipFieldColumnHidden_notAdded() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("columnHidden", "true");
        obj.addProperty("address", "A1");
        FreeForm ff = invokeGraphObjectToFreeForm(obj);
        assertNull(ff.get("columnHidden"), "RANGE_SKIP_FIELDS must be excluded from FreeForm");
        assertNotNull(ff.get("Cell"));
    }

    @Test
    void graphObjectToFreeForm_rangeSkipFieldRowHidden_notAdded() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("rowHidden", "false");
        FreeForm ff = invokeGraphObjectToFreeForm(obj);
        assertNull(ff.get("rowHidden"));
    }

    @Test
    void graphObjectToFreeForm_rangeSkipFieldNumberFormat_notAdded() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("numberFormat", "General");
        obj.addProperty("name", "Sheet1");
        FreeForm ff = invokeGraphObjectToFreeForm(obj);
        assertNull(ff.get("numberFormat"));
    }

    @Test
    void graphObjectToFreeForm_emptyStringValue_notAddedToFreeForm() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("name", "File.xlsx");
        obj.addProperty("webUrl", ""); // toDisplayString returns "" → should be omitted
        FreeForm ff = invokeGraphObjectToFreeForm(obj);
        assertNotNull(ff.get("Name"));
        // empty webUrl → Link column must not be present
        assertNull(ff.get("Link"), "Empty string values must not be added as FreeForm entries");
    }

    // ── toDisplayString — additional branches ─────────────────────────────────

    @Test
    void toDisplayString_keyContainsLowercaseDate_withT_truncatesToDate() throws Exception {
        // key "somedate" contains lowercase "date" → s.contains("T") → truncated
        String result = invokeToDisplayString("somedate",
                new JsonPrimitive("2024-06-01T09:00:00Z"), new JsonObject());
        assertEquals("2024-06-01", result);
    }

    @Test
    void toDisplayString_keyContainsLowercaseDate_noT_returnedAsIs() throws Exception {
        // value has no "T" → condition false → returned as-is
        String result = invokeToDisplayString("somedate",
                new JsonPrimitive("2024-06-01"), new JsonObject());
        assertEquals("2024-06-01", result);
    }

    @Test
    void toDisplayString_sizeNonNumeric_returnsAsIs() throws Exception {
        // NumberFormatException catch branch
        String result = invokeToDisplayString("size",
                new JsonPrimitive("unknown-size"), new JsonObject());
        assertEquals("unknown-size", result);
    }

    @Test
    void toDisplayString_lastModifiedBy_noUserKey_returnsEmpty() throws Exception {
        // identitySet without "user" key
        JsonObject identitySet = new JsonObject();
        identitySet.addProperty("application", "SomeApp");
        String result = invokeToDisplayString("lastModifiedBy", identitySet, new JsonObject());
        assertEquals("", result);
    }

    @Test
    void toDisplayString_lastModifiedBy_nullUser_returnsEmpty() throws Exception {
        JsonObject identitySet = new JsonObject();
        identitySet.add("user", JsonNull.INSTANCE);
        String result = invokeToDisplayString("lastModifiedBy", identitySet, new JsonObject());
        assertEquals("", result);
    }

    @Test
    void toDisplayString_valuesArray_withNonArrayRow_skipsNonArrayRow() throws Exception {
        // row element is not a JsonArray → `if (!rowEl.isJsonArray()) continue;`
        JsonArray arr = new JsonArray();
        arr.add(new JsonPrimitive("not-an-array-row")); // non-array → skipped
        JsonArray realRow = new JsonArray();
        realRow.add("A1");
        arr.add(realRow);

        String result = invokeToDisplayString("values", arr, new JsonObject());
        // Only the real row is included
        assertEquals("A1", result);
    }

    @Test
    void toDisplayString_valuesArray_withNullCell_rendersSecondCellOnly() throws Exception {
        // null first cell → empty string; next cell "B1" has no separator prefix
        JsonArray row = new JsonArray();
        row.add(JsonNull.INSTANCE); // null cell → ""
        row.add("B1");
        JsonArray arr = new JsonArray();
        arr.add(row);

        String result = invokeToDisplayString("values", arr, new JsonObject());
        assertNotNull(result);
        assertTrue(result.contains("B1"));
    }

    // ── flattenValueArray — additional branches ───────────────────────────────

    @Test
    void flattenValueArray_nullInput_returnsNull() throws Exception {
        assertNull(invokeFlattenValueArray(null));
    }

    @Test
    void flattenValueArray_rowWithNullCell_nullRenderedAsEmptyNoSeparatorBefore() throws Exception {
        // null cell → "" (empty), next cell "Val" appended without leading " | "
        // because rowSb was empty after the null cell
        JsonArray row = new JsonArray();
        row.add(JsonNull.INSTANCE);
        row.add("Val");
        JsonArray arr = new JsonArray();
        arr.add(row);
        String result = invokeFlattenValueArray(arr);
        assertNotNull(result);
        assertTrue(result.contains("Val"));
        // No leading separator since null cell contributes empty string
        assertFalse(result.startsWith("|"));
    }

    @Test
    void flattenValueArray_allNullCells_rowSkipped() throws Exception {
        // All cells null → rowSb stays empty → row not appended
        JsonArray row = new JsonArray();
        row.add(JsonNull.INSTANCE);
        JsonArray arr = new JsonArray();
        arr.add(row);
        // null cell → rowSb = "" → rowSb.length() == 0 → not appended
        // sb stays empty → returns null
        assertNull(invokeFlattenValueArray(arr));
    }

    // ── parseGenericResponse — additional branches ────────────────────────────

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_flatListWithNonObjectElement_skipsIt() throws Exception {
        // `if (!el.isJsonObject()) continue;` branch
        String json = "{\"value\":[\"not-an-object\",{\"name\":\"Real.xlsx\"}]}";
        List<FreeForm> rows = invokeParseGenericResponse(json);
        assertEquals(1, rows.size(), "Non-object elements in flat list must be skipped");
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_searchWithNoHitsContainersKey_returnsEmpty() throws Exception {
        // First element has hitsContainers but empty hitsContainers array
        String json = "{\"value\":[{\"hitsContainers\":[]}]}";
        List<FreeForm> rows = invokeParseGenericResponse(json);
        assertTrue(rows.isEmpty());
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_searchWithEmptyHitsArray_returnsEmpty() throws Exception {
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[]}]}]}";
        List<FreeForm> rows = invokeParseGenericResponse(json);
        assertTrue(rows.isEmpty());
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_searchHitWithoutHitsKey_returnsEmpty() throws Exception {
        // hitsContainer has no "hits" key
        String json = "{\"value\":[{\"hitsContainers\":[{}]}]}";
        List<FreeForm> rows = invokeParseGenericResponse(json);
        assertTrue(rows.isEmpty());
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_searchHitWithNullHitId_fileIdFromResource() throws Exception {
        // hitId present but null → falls through; resource.id used via graphObjectToFreeForm
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":null,"
                + "\"resource\":{\"name\":\"File.xlsx\",\"id\":\"res-001\"}"
                + "}]}]}]}";
        List<FreeForm> rows = invokeParseGenericResponse(json);
        assertEquals(1, rows.size());
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_searchHitWithEmptyHitId_notAddedAsFileId() throws Exception {
        // hitId is empty string → `if (!hitId.isEmpty())` branch → FileId from graphObjectToFreeForm
        String json = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":\"\","
                + "\"resource\":{\"name\":\"File.xlsx\"}"
                + "}]}]}]}";
        List<FreeForm> rows = invokeParseGenericResponse(json);
        assertEquals(1, rows.size());
    }

    @Test
    @SuppressWarnings("unchecked")
    void parseGenericResponse_flatList_nonObjectFollowedByObject() throws Exception {
        // Mix: primitive then object
        String json = "{\"value\":[42,{\"name\":\"Book.xlsx\",\"id\":\"id1\"}]}";
        List<FreeForm> rows = invokeParseGenericResponse(json);
        assertEquals(1, rows.size());
    }

    // ── getJsonString ─────────────────────────────────────────────────────────

    @Test
    void getJsonString_nullObject_returnsEmpty() throws Exception {
        assertEquals("", invokeGetJsonString(null, "key"));
    }

    @Test
    void getJsonString_missingKey_returnsEmpty() throws Exception {
        assertEquals("", invokeGetJsonString(new JsonObject(), "missing"));
    }

    @Test
    void getJsonString_nullValue_returnsEmpty() throws Exception {
        JsonObject obj = new JsonObject();
        obj.add("key", JsonNull.INSTANCE);
        assertEquals("", invokeGetJsonString(obj, "key"));
    }

    @Test
    void getJsonString_presentValue_returnsString() throws Exception {
        JsonObject obj = new JsonObject();
        obj.addProperty("method", "GET");
        assertEquals("GET", invokeGetJsonString(obj, "method"));
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
        when(refreshTokenStore.get("k1")).thenReturn("myToken");
        ITokenStore store = invokeAdaptTokenStore(refreshTokenStore);
        assertEquals("myToken", store.get("k1"));
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

    // ── Reflection helpers ────────────────────────────────────────────────────

    private static FreeForm invokeGraphObjectToFreeForm(JsonObject obj) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("graphObjectToFreeForm", JsonObject.class);
        m.setAccessible(true);
        return (FreeForm) m.invoke(null, obj);
    }

    private static String invokeToDisplayString(String key,
                                                com.google.gson.JsonElement val,
                                                JsonObject parent) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "toDisplayString", String.class,
                com.google.gson.JsonElement.class, JsonObject.class);
        m.setAccessible(true);
        return (String) m.invoke(null, key, val, parent);
    }

    private static String invokeFlattenValueArray(JsonArray arr) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("flattenValueArray", JsonArray.class);
        m.setAccessible(true);
        return (String) m.invoke(null, arr);
    }

    @SuppressWarnings("unchecked")
    private static List<FreeForm> invokeParseGenericResponse(String json) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod("parseGenericResponse", String.class);
        m.setAccessible(true);
        return (List<FreeForm>) m.invoke(null, json);
    }

    private static String invokeGetJsonString(JsonObject obj, String key) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "getJsonString", JsonObject.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, obj, key);
    }

    private static ITokenStore invokeAdaptTokenStore(RefreshTokenStore store) throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "adaptTokenStore", RefreshTokenStore.class);
        m.setAccessible(true);
        return (ITokenStore) m.invoke(null, store);
    }
}
