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
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.model.base.FreeForm;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.lenient;

/**
 * Unit tests for the private helper methods in {@link ExcelMcpSetupResource}.
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>{@code convertToNlqResponse} — dynamic column derivation, row mapping, generatedQuery</li>
 *   <li>{@code toColumnKey} — PascalCase → camelCase normalisation</li>
 *   <li>{@code toColumnLabel} — camelCase → human-readable label</li>
 *   <li>{@code isLinkColumn} — identifies link-type columns (name, link, url)</li>
 *   <li>{@code extractFreeFormValue} — FreeForm [type, value] list extraction</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
class ExcelMcpSetupResourceTest {

    @Mock private RefreshTokenStore         refreshTokenStore;
    @Mock private Invoker                   invoker;
    @Mock private OfficeExcelAttributeStore attributeStore;
    @Mock private ExcelMcpQueryService      queryService;

    private ExcelMcpSetupResource resource;

    // Reflected methods
    private Method convertToNlqResponse;
    private Method toColumnKey;
    private Method toColumnLabel;
    private Method isLinkColumn;
    private Method extractFreeFormValue;

    @BeforeEach
    void setUp() throws Exception {
        lenient().when(invoker.getInvokerId()).thenReturn("test-invoker");
        lenient().when(attributeStore.load("test-invoker")).thenReturn(null);
        lenient().when(refreshTokenStore.get(anyString())).thenReturn(null);

        resource = new ExcelMcpSetupResource(
                refreshTokenStore, invoker, attributeStore, queryService);

        convertToNlqResponse = ExcelMcpSetupResource.class
                .getDeclaredMethod("convertToNlqResponse",
                        ExtensionResponse.class, String.class);
        convertToNlqResponse.setAccessible(true);

        toColumnKey = ExcelMcpSetupResource.class
                .getDeclaredMethod("toColumnKey", String.class);
        toColumnKey.setAccessible(true);

        toColumnLabel = ExcelMcpSetupResource.class
                .getDeclaredMethod("toColumnLabel", String.class);
        toColumnLabel.setAccessible(true);

        isLinkColumn = ExcelMcpSetupResource.class
                .getDeclaredMethod("isLinkColumn", String.class);
        isLinkColumn.setAccessible(true);

        extractFreeFormValue = ExcelMcpSetupResource.class
                .getDeclaredMethod("extractFreeFormValue", FreeForm.class, Object.class);
        extractFreeFormValue.setAccessible(true);
    }

    // ── convertToNlqResponse — basic structure ────────────────────────────

    @Test
    void convertToNlqResponse_emptyRows_returnsTotalZero() throws Exception {
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of());
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        assertEquals(0, result.get("total"));
        assertEquals(0, result.get("shown"));
    }

    @Test
    void convertToNlqResponse_alwaysHasSpecificFalse() throws Exception {
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of());
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        assertEquals(Boolean.FALSE, result.get("specific"));
    }

    @Test
    void convertToNlqResponse_generatedQueryDefaultsToOriginalQuestion() throws Exception {
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of());
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "list my Excel files");
        assertEquals("list my Excel files", result.get("generatedQuery"),
                "Without generatedPlan, generatedQuery must default to original question");
    }

    @Test
    void convertToNlqResponse_withGeneratedPlan_usesGraphPlan() throws Exception {
        Map<String, Object> payload = new HashMap<>();
        payload.put("generatedPlan", "GET /me/drive/root/children");
        payload.put("Response", List.of());
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "list files");
        assertEquals("GET /me/drive/root/children", result.get("generatedQuery"),
                "generatedQuery must show the AI-generated Graph plan, not the original question");
    }

    @Test
    void convertToNlqResponse_searchGeneratedPlan_usesSearchPath() throws Exception {
        String searchPlan = "POST /search/query";
        Map<String, Object> payload = new HashMap<>();
        payload.put("generatedPlan", searchPlan);
        payload.put("Response", List.of());
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "find budget files");
        assertEquals(searchPlan, result.get("generatedQuery"));
    }

    @Test
    void convertToNlqResponse_nullPayload_returnsTotalZero() throws Exception {
        ExtensionResponse resp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, null, null, null, null);
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, resp, "q");
        assertEquals(0, result.get("total"));
    }

    @Test
    void convertToNlqResponse_nullPayload_generatedQueryIsOriginalQuestion() throws Exception {
        ExtensionResponse resp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, null, null, null, null);
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, resp, "original question");
        assertEquals("original question", result.get("generatedQuery"),
                "Null response value must fall back to original question");
    }

    // ── convertToNlqResponse — row mapping ────────────────────────────────

    @Test
    void convertToNlqResponse_singleRow_returnsCorrectCount() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name", " Text ", "Budget.xlsx");
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of(ff));
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        assertEquals(1, result.get("total"));
        assertEquals(1, result.get("shown"));
    }

    @Test
    void convertToNlqResponse_freeFormRow_columnKeyIsCamelCase() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name", " Text ", "Budget.xlsx");
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of(ff));
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        @SuppressWarnings("unchecked")
        List<Map<String, String>> rows = (List<Map<String, String>>) result.get("rows");
        assertEquals("Budget.xlsx", rows.get(0).get("name"),
                "PascalCase FreeForm key 'Name' must become camelCase 'name'");
    }

    @Test
    void convertToNlqResponse_multipleRows_allReturned() throws Exception {
        FreeForm ff1 = new FreeForm();
        ff1.put("Name", " Text ", "File1.xlsx");
        FreeForm ff2 = new FreeForm();
        ff2.put("Name", " Text ", "File2.xlsx");
        FreeForm ff3 = new FreeForm();
        ff3.put("Name", " Text ", "File3.xlsx");
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of(ff1, ff2, ff3));
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        assertEquals(3, result.get("total"));
    }

    @Test
    void convertToNlqResponse_emptyFreeFormValue_columnOmitted() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name", " Text ", "Budget.xlsx");
        ff.put("Size", " Text ", ""); // empty value should be omitted
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of(ff));
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        @SuppressWarnings("unchecked")
        List<Map<String, String>> rows = (List<Map<String, String>>) result.get("rows");
        assertFalse(rows.get(0).containsKey("size"),
                "Empty FreeForm values must not produce a column entry");
    }

    // ── convertToNlqResponse — column definitions ─────────────────────────

    @Test
    void convertToNlqResponse_nameColumn_isTypedAsLink() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name", " Text ", "Budget.xlsx");
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of(ff));
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        @SuppressWarnings("unchecked")
        List<Map<String, String>> columns = (List<Map<String, String>>) result.get("columns");
        Map<String, String> nameCol = columns.stream()
                .filter(c -> "name".equals(c.get("key"))).findFirst().orElse(null);
        assertNotNull(nameCol, "name column must exist");
        assertEquals("link", nameCol.get("type"),
                "name column must be type 'link' so the UI renders it as a clickable hyperlink");
    }

    @Test
    void convertToNlqResponse_lastModifiedColumn_isTypedAsText() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("LastModified", " Text ", "2024-05-15");
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of(ff));
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        @SuppressWarnings("unchecked")
        List<Map<String, String>> columns = (List<Map<String, String>>) result.get("columns");
        Map<String, String> col = columns.stream()
                .filter(c -> "lastModified".equals(c.get("key"))).findFirst().orElse(null);
        assertNotNull(col, "lastModified column must exist");
        assertEquals("text", col.get("type"),
                "lastModified column must be type 'text' (not a clickable link)");
    }

    @Test
    void convertToNlqResponse_columnsAreDerivedDynamically() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name",         " Text ", "Budget.xlsx");
        ff.put("LastModified", " Text ", "2024-05-15");
        ff.put("ModifiedBy",   " Text ", "Alice Smith");
        Map<String, Object> payload = new HashMap<>();
        payload.put("Response", List.of(ff));
        @SuppressWarnings("unchecked")
        Map<String, Object> result = (Map<String, Object>)
                convertToNlqResponse.invoke(resource, successResp(payload), "q");
        @SuppressWarnings("unchecked")
        List<Map<String, String>> columns = (List<Map<String, String>>) result.get("columns");
        List<String> colKeys = columns.stream().map(c -> c.get("key")).toList();
        assertTrue(colKeys.contains("name"),          "name must be in derived columns");
        assertTrue(colKeys.contains("lastModified"),  "lastModified must be in derived columns");
        assertTrue(colKeys.contains("modifiedBy"),    "modifiedBy must be in derived columns");
    }

    // ── toColumnKey ────────────────────────────────────────────────────────

    @Test
    void toColumnKey_pascalCase_lowercasesFirstChar() throws Exception {
        assertEquals("name",         toColumnKey.invoke(null, "Name"));
        assertEquals("lastModified", toColumnKey.invoke(null, "LastModified"));
        assertEquals("fileId",       toColumnKey.invoke(null, "FileId"));
        assertEquals("modifiedBy",   toColumnKey.invoke(null, "ModifiedBy"));
        assertEquals("createdBy",    toColumnKey.invoke(null, "CreatedBy"));
    }

    @Test
    void toColumnKey_alreadyCamelCase_unchanged() throws Exception {
        assertEquals("name", toColumnKey.invoke(null, "name"));
    }

    @Test
    void toColumnKey_emptyString_returnsEmpty() throws Exception {
        assertEquals("", toColumnKey.invoke(null, ""));
    }

    @Test
    void toColumnKey_null_returnsNull() throws Exception {
        assertNull(toColumnKey.invoke(null, (Object) null));
    }

    // ── toColumnLabel ──────────────────────────────────────────────────────

    @Test
    void toColumnLabel_camelCase_insertsSpacesBeforeUppercase() throws Exception {
        assertEquals("Last Modified", toColumnLabel.invoke(null, "lastModified"));
        assertEquals("Modified By",  toColumnLabel.invoke(null, "modifiedBy"));
        assertEquals("File Id",      toColumnLabel.invoke(null, "fileId"));
        assertEquals("Created By",   toColumnLabel.invoke(null, "createdBy"));
    }

    @Test
    void toColumnLabel_singleWord_capitalisesFirst() throws Exception {
        assertEquals("Name", toColumnLabel.invoke(null, "name"));
        assertEquals("Size", toColumnLabel.invoke(null, "size"));
        assertEquals("Link", toColumnLabel.invoke(null, "link"));
    }

    @Test
    void toColumnLabel_emptyString_returnsEmpty() throws Exception {
        assertEquals("", toColumnLabel.invoke(null, ""));
    }

    @Test
    void toColumnLabel_null_returnsNull() throws Exception {
        assertNull(toColumnLabel.invoke(null, (Object) null));
    }

    // ── isLinkColumn ────────────────────────────────────────────────────────

    @Test
    void isLinkColumn_name_isTrue() throws Exception {
        assertTrue((boolean) isLinkColumn.invoke(null, "name"),
                "'name' is the file title column and must be rendered as a link");
    }

    @Test
    void isLinkColumn_link_isTrue() throws Exception {
        assertTrue((boolean) isLinkColumn.invoke(null, "link"));
    }

    @Test
    void isLinkColumn_url_isTrue() throws Exception {
        assertTrue((boolean) isLinkColumn.invoke(null, "url"));
    }

    @Test
    void isLinkColumn_fileId_isFalse() throws Exception {
        assertFalse((boolean) isLinkColumn.invoke(null, "fileId"));
    }

    @Test
    void isLinkColumn_lastModified_isFalse() throws Exception {
        assertFalse((boolean) isLinkColumn.invoke(null, "lastModified"));
    }

    @Test
    void isLinkColumn_size_isFalse() throws Exception {
        assertFalse((boolean) isLinkColumn.invoke(null, "size"));
    }

    @Test
    void isLinkColumn_modifiedBy_isFalse() throws Exception {
        assertFalse((boolean) isLinkColumn.invoke(null, "modifiedBy"));
    }

    // ── extractFreeFormValue ───────────────────────────────────────────────

    @Test
    void extractFreeFormValue_standardEntry_returnsValue() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name", " Text ", "Budget.xlsx");
        assertEquals("Budget.xlsx",
                extractFreeFormValue.invoke(resource, ff, "Name"));
    }

    @Test
    void extractFreeFormValue_emptyValue_returnsNull() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name", " Text ", "");
        assertNull(extractFreeFormValue.invoke(resource, ff, "Name"),
                "Empty string values must be normalised to null");
    }

    @Test
    void extractFreeFormValue_missingKey_returnsNull() throws Exception {
        FreeForm ff = new FreeForm();
        assertNull(extractFreeFormValue.invoke(resource, ff, "Name"),
                "Missing FreeForm key must return null");
    }

    @Test
    void extractFreeFormValue_multipleFields_eachExtractedCorrectly() throws Exception {
        FreeForm ff = new FreeForm();
        ff.put("Name",         " Text ", "Budget.xlsx");
        ff.put("LastModified", " Text ", "2024-05-15");
        ff.put("Size",         " Text ", "2 KB");
        assertEquals("Budget.xlsx", extractFreeFormValue.invoke(resource, ff, "Name"));
        assertEquals("2024-05-15",  extractFreeFormValue.invoke(resource, ff, "LastModified"));
        assertEquals("2 KB",        extractFreeFormValue.invoke(resource, ff, "Size"));
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    private static ExtensionResponse successResp(Map<String, Object> payload) {
        return new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
    }
}
