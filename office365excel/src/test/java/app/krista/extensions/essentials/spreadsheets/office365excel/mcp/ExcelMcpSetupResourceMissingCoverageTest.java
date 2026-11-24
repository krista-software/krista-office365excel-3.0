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
import app.krista.sdk.msauth.extension.mcp.McpProductConfig;
import app.krista.sdk.msauth.token.ITokenStore;
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
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.*;

/**
 * Supplemental unit tests for {@link ExcelMcpSetupResource} covering branches
 * not exercised by {@link ExcelMcpSetupResourceTest}.
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>{@code getProductConfig()} — returns an {@link ExcelMcpProductConfig} instance</li>
 *   <li>{@code getJaxrsId()} — returns {@code "excel"}</li>
 *   <li>{@code interceptNlQuery()} — FAILURE (non-null error), FAILURE (null error),
 *       SUCCESS, and Exception paths</li>
 *   <li>{@code adaptTokenStore()} — save/get/delete delegation and IOException wrapping</li>
 *   <li>{@code extractFreeFormValue()} — non-List val branch, List size &lt; 2,
 *       List with {@code null} at index 1</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class ExcelMcpSetupResourceMissingCoverageTest {

    @Mock private RefreshTokenStore         refreshTokenStore;
    @Mock private Invoker                   invoker;
    @Mock private OfficeExcelAttributeStore attributeStore;
    @Mock private ExcelMcpQueryService      queryService;

    private ExcelMcpSetupResource resource;

    // Reflected methods
    private Method interceptNlQuery;
    private Method extractFreeFormValue;
    private Method adaptTokenStore;
    private Method getProductConfig;
    private Method getJaxrsId;

    @BeforeEach
    void setUp() throws Exception {
        lenient().when(invoker.getInvokerId()).thenReturn("test-invoker");
        lenient().when(attributeStore.load("test-invoker")).thenReturn(null);
        lenient().when(refreshTokenStore.get(anyString())).thenReturn(null);

        resource = new ExcelMcpSetupResource(
                refreshTokenStore, invoker, attributeStore, queryService);

        interceptNlQuery = ExcelMcpSetupResource.class
                .getDeclaredMethod("interceptNlQuery", String.class);
        interceptNlQuery.setAccessible(true);

        extractFreeFormValue = ExcelMcpSetupResource.class
                .getDeclaredMethod("extractFreeFormValue", FreeForm.class, Object.class);
        extractFreeFormValue.setAccessible(true);

        adaptTokenStore = ExcelMcpSetupResource.class
                .getDeclaredMethod("adaptTokenStore", RefreshTokenStore.class);
        adaptTokenStore.setAccessible(true);

        getProductConfig = ExcelMcpSetupResource.class
                .getDeclaredMethod("getProductConfig");
        getProductConfig.setAccessible(true);

        getJaxrsId = ExcelMcpSetupResource.class
                .getDeclaredMethod("getJaxrsId");
        getJaxrsId.setAccessible(true);
    }

    // ── getProductConfig ────────────────────────────────────────────────────

    @Test
    void getProductConfig_returnsExcelMcpProductConfigInstance() throws Exception {
        McpProductConfig config = (McpProductConfig) getProductConfig.invoke(resource);
        assertInstanceOf(ExcelMcpProductConfig.class, config,
                "getProductConfig() must return an ExcelMcpProductConfig instance");
    }

    @Test
    void getProductConfig_returnsNonNullInstance() throws Exception {
        assertNotNull(getProductConfig.invoke(resource));
    }

    // ── getJaxrsId ──────────────────────────────────────────────────────────

    @Test
    void getJaxrsId_returnsExcel() throws Exception {
        String jaxrsId = (String) getJaxrsId.invoke(resource);
        assertEquals("excel", jaxrsId, "getJaxrsId() must return \"excel\"");
    }

    // ── interceptNlQuery — FAILURE with non-null error ──────────────────────

    @Test
    void interceptNlQuery_failureWithNonNullError_returns400() throws Exception {
        ExtensionResponse.Error error = new ExtensionResponse.Error(
                "Token expired", System.currentTimeMillis(),
                ExtensionResponse.Error.ExceptionType.AUTHENTICATION_ERROR, null);
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, error, null, null);
        when(queryService.query("test question")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "test question");

        assertEquals(400, response.getStatus());
    }

    @Test
    void interceptNlQuery_failureWithNonNullError_bodyContainsErrorKey() throws Exception {
        ExtensionResponse.Error error = new ExtensionResponse.Error(
                "Token expired", System.currentTimeMillis(),
                ExtensionResponse.Error.ExceptionType.AUTHENTICATION_ERROR, null);
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, error, null, null);
        when(queryService.query("q")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        assertTrue(response.getEntity().toString().contains("\"error\""),
                "FAILURE response must contain an 'error' JSON key");
    }

    @Test
    void interceptNlQuery_failureWithNonNullError_bodyIsNotDefaultMessage() throws Exception {
        // When error is non-null the body must NOT fall back to "Query failed."
        ExtensionResponse.Error error = new ExtensionResponse.Error(
                "Auth error", System.currentTimeMillis(),
                ExtensionResponse.Error.ExceptionType.AUTHENTICATION_ERROR, null);
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, error, null, null);
        when(queryService.query("q")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        // Verify error message comes from the Error object (via String.valueOf), not the fallback
        assertFalse(response.getEntity().toString().equals("{\"error\":\"Query failed.\"}"),
                "When error is non-null, error message must come from extResponse.getError()");
    }

    @Test
    void interceptNlQuery_failureWithNonNullError_contentTypeIsJson() throws Exception {
        ExtensionResponse.Error error = new ExtensionResponse.Error(
                "Some error", System.currentTimeMillis(),
                ExtensionResponse.Error.ExceptionType.SYSTEM_ERROR, null);
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, error, null, null);
        when(queryService.query("q")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        assertTrue(response.getMediaType().toString().contains("application/json"),
                "Error response must have Content-Type application/json");
    }

    // ── interceptNlQuery — FAILURE with null error ──────────────────────────

    @Test
    void interceptNlQuery_failureWithNullError_returns400() throws Exception {
        // null as Error (3rd param) → getError() returns null → default "Query failed." message
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, null, null, null);
        when(queryService.query("test question")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "test question");

        assertEquals(400, response.getStatus());
    }

    @Test
    void interceptNlQuery_failureWithNullError_bodyContainsDefaultMessage() throws Exception {
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, null, null, null);
        when(queryService.query("test question")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "test question");

        String entity = response.getEntity().toString();
        assertTrue(entity.contains("Query failed."),
                "When error is null, response must use default message 'Query failed.'");
    }

    @Test
    void interceptNlQuery_failureWithNullError_bodyContainsErrorKey() throws Exception {
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, null, null, null);
        when(queryService.query("q")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        assertTrue(response.getEntity().toString().contains("\"error\""),
                "JSON response must include an 'error' key");
    }

    @Test
    void interceptNlQuery_failureWithNullError_contentTypeIsJson() throws Exception {
        ExtensionResponse failResp = new ExtensionResponse(
                ExtensionResponse.Result.FAILURE, null, null, null, null);
        when(queryService.query("q")).thenReturn(failResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        assertTrue(response.getMediaType().toString().contains("application/json"));
    }

    // ── interceptNlQuery — SUCCESS ──────────────────────────────────────────

    @Test
    void interceptNlQuery_success_returns200() throws Exception {
        Map<String, Object> payload = Map.of("Response", List.of());
        ExtensionResponse successResp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
        when(queryService.query("list files")).thenReturn(successResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "list files");

        assertEquals(200, response.getStatus());
    }

    @Test
    void interceptNlQuery_success_bodyContainsTotalField() throws Exception {
        Map<String, Object> payload = Map.of("Response", List.of());
        ExtensionResponse successResp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
        when(queryService.query("list files")).thenReturn(successResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "list files");

        assertTrue(response.getEntity().toString().contains("\"total\""),
                "Successful 200 response must contain 'total' field");
    }

    @Test
    void interceptNlQuery_success_bodyContainsGeneratedQuery() throws Exception {
        Map<String, Object> payload = Map.of("Response", List.of());
        ExtensionResponse successResp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
        when(queryService.query("list my files")).thenReturn(successResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "list my files");

        assertTrue(response.getEntity().toString().contains("\"generatedQuery\""),
                "Successful 200 response must contain 'generatedQuery' field");
    }

    @Test
    void interceptNlQuery_success_contentTypeIsJson() throws Exception {
        Map<String, Object> payload = Map.of("Response", List.of());
        ExtensionResponse successResp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
        when(queryService.query("q")).thenReturn(successResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        assertTrue(response.getMediaType().toString().contains("application/json"),
                "Successful response must have Content-Type application/json");
    }

    @Test
    void interceptNlQuery_success_withGeneratedPlan_reflectedInBody() throws Exception {
        Map<String, Object> payload = new HashMap<>();
        payload.put("generatedPlan", "GET /me/drive/root/children");
        payload.put("Response", List.of());
        ExtensionResponse successResp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
        when(queryService.query("list files")).thenReturn(successResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "list files");

        assertTrue(response.getEntity().toString().contains("GET /me/drive/root/children"),
                "generatedQuery in response must reflect the AI-generated plan");
    }

    @Test
    void interceptNlQuery_success_bodyContainsRowsField() throws Exception {
        Map<String, Object> payload = Map.of("Response", List.of());
        ExtensionResponse successResp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
        when(queryService.query("q")).thenReturn(successResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        assertTrue(response.getEntity().toString().contains("\"rows\""));
    }

    @Test
    void interceptNlQuery_success_bodyContainsColumnsField() throws Exception {
        Map<String, Object> payload = Map.of("Response", List.of());
        ExtensionResponse successResp = new ExtensionResponse(
                ExtensionResponse.Result.SUCCESS, payload, null, null, null);
        when(queryService.query("q")).thenReturn(successResp);

        Response response = (Response) interceptNlQuery.invoke(resource, "q");

        assertTrue(response.getEntity().toString().contains("\"columns\""));
    }

    // ── interceptNlQuery — Exception path ──────────────────────────────────

    @Test
    void interceptNlQuery_exceptionFromQueryService_returns400() throws Exception {
        when(queryService.query("crash")).thenThrow(new RuntimeException("connection refused"));

        Response response = (Response) interceptNlQuery.invoke(resource, "crash");

        assertEquals(400, response.getStatus());
    }

    @Test
    void interceptNlQuery_exceptionFromQueryService_bodyContainsQueryFailed() throws Exception {
        when(queryService.query("crash")).thenThrow(new RuntimeException("connection refused"));

        Response response = (Response) interceptNlQuery.invoke(resource, "crash");

        assertTrue(response.getEntity().toString().contains("Query failed:"),
                "Exception path must return 'Query failed: ...' message");
    }

    @Test
    void interceptNlQuery_exceptionFromQueryService_bodyContainsExceptionMessage() throws Exception {
        when(queryService.query("crash")).thenThrow(new RuntimeException("connection refused"));

        Response response = (Response) interceptNlQuery.invoke(resource, "crash");

        assertTrue(response.getEntity().toString().contains("connection refused"),
                "Exception path must include the original exception message");
    }

    @Test
    void interceptNlQuery_exceptionFromQueryService_contentTypeIsJson() throws Exception {
        when(queryService.query("crash")).thenThrow(new RuntimeException("timeout"));

        Response response = (Response) interceptNlQuery.invoke(resource, "crash");

        assertTrue(response.getMediaType().toString().contains("application/json"));
    }

    @Test
    void interceptNlQuery_exceptionFromQueryService_bodyContainsErrorKey() throws Exception {
        when(queryService.query("crash")).thenThrow(new RuntimeException("timeout"));

        Response response = (Response) interceptNlQuery.invoke(resource, "crash");

        assertTrue(response.getEntity().toString().contains("\"error\""));
    }

    // ── adaptTokenStore — save delegation ──────────────────────────────────

    @Test
    void adaptTokenStore_save_delegatesToRefreshTokenStorePut() throws Exception {
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        tokenStore.save("my-key", "my-value");

        verify(refreshTokenStore).put("my-key", "my-value");
    }

    @Test
    void adaptTokenStore_save_ioExceptionWrappedAsUnchecked() throws Exception {
        doThrow(new IOException("disk full")).when(refreshTokenStore).put(anyString(), anyString());
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        assertThrows(UncheckedIOException.class, () -> tokenStore.save("k", "v"),
                "IOException from RefreshTokenStore.put must be wrapped as UncheckedIOException");
    }

    @Test
    void adaptTokenStore_save_ioExceptionCausePreserved() throws Exception {
        IOException cause = new IOException("disk full");
        doThrow(cause).when(refreshTokenStore).put(anyString(), anyString());
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        UncheckedIOException thrown = assertThrows(UncheckedIOException.class,
                () -> tokenStore.save("k", "v"));
        assertSame(cause, thrown.getCause(),
                "UncheckedIOException must wrap the original IOException as cause");
    }

    // ── adaptTokenStore — get delegation ───────────────────────────────────

    @Test
    void adaptTokenStore_get_delegatesToRefreshTokenStoreGet() throws Exception {
        when(refreshTokenStore.get("my-key")).thenReturn("my-token");
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        String result = tokenStore.get("my-key");

        assertEquals("my-token", result);
    }

    @Test
    void adaptTokenStore_get_nullValuePropagates() throws Exception {
        when(refreshTokenStore.get("missing")).thenReturn(null);
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        assertNull(tokenStore.get("missing"),
                "Null token must propagate from RefreshTokenStore.get()");
    }

    @Test
    void adaptTokenStore_get_ioExceptionWrappedAsUnchecked() throws Exception {
        when(refreshTokenStore.get(anyString())).thenThrow(new IOException("read error"));
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        assertThrows(UncheckedIOException.class, () -> tokenStore.get("k"),
                "IOException from RefreshTokenStore.get must be wrapped as UncheckedIOException");
    }

    @Test
    void adaptTokenStore_get_ioExceptionCausePreserved() throws Exception {
        IOException cause = new IOException("read error");
        when(refreshTokenStore.get(anyString())).thenThrow(cause);
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        UncheckedIOException thrown = assertThrows(UncheckedIOException.class,
                () -> tokenStore.get("k"));
        assertSame(cause, thrown.getCause());
    }

    // ── adaptTokenStore — delete delegation ────────────────────────────────

    @Test
    void adaptTokenStore_delete_delegatesToRefreshTokenStoreRemove() throws Exception {
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        tokenStore.delete("my-key");

        verify(refreshTokenStore).remove("my-key");
    }

    @Test
    void adaptTokenStore_delete_doesNotThrowForMissingKey() throws Exception {
        ITokenStore tokenStore = (ITokenStore) adaptTokenStore.invoke(null, refreshTokenStore);

        assertDoesNotThrow(() -> tokenStore.delete("non-existent-key"));
    }

    // ── extractFreeFormValue — val instanceof List branch ─────────────────
    //
    // FreeForm.get() returns typedValue.getValue() directly.
    // To exercise the "val instanceof List" branch we must store a TypedValue
    // whose wrapped value IS a List.  FreeForm.put(String, Object) only accepts
    // FreeForm$TypedValue instances, so we create them via reflection.

    /** Creates a FreeForm$TypedValue wrapping an arbitrary value using reflection. */
    private static Object createTypedValue(Object value) throws Exception {
        Class<?> tvClass = Class.forName("app.krista.model.base.FreeForm$TypedValue");
        // 2-arg constructor: TypedValue(String typeDescription, Object value)
        java.lang.reflect.Constructor<?> ctor =
                tvClass.getDeclaredConstructor(String.class, Object.class);
        ctor.setAccessible(true);
        return ctor.newInstance(" Text ", value);
    }

    @Test
    void extractFreeFormValue_valueIsListSizeGe2NonNullIndex1_returnsValue() throws Exception {
        // val = ["type", "Budget.xlsx"] → size 2, get(1) = "Budget.xlsx"
        FreeForm ff = new FreeForm();
        ff.put("Name", createTypedValue(Arrays.asList("type", "Budget.xlsx")));

        String result = (String) extractFreeFormValue.invoke(resource, ff, "Name");

        assertEquals("Budget.xlsx", result,
                "List with size >= 2 and non-null index-1 must return the index-1 value");
    }

    @Test
    void extractFreeFormValue_valueIsListSizeGe2EmptyStringIndex1_returnsNull() throws Exception {
        // val = ["type", ""] → size 2, get(1) = "" → s.isEmpty() → null
        FreeForm ff = new FreeForm();
        ff.put("Size", createTypedValue(Arrays.asList("type", "")));

        String result = (String) extractFreeFormValue.invoke(resource, ff, "Size");

        assertNull(result,
                "List with empty string at index 1 must be normalised to null");
    }

    @Test
    void extractFreeFormValue_valueIsListWithNullAtIndex1_returnsNull() throws Exception {
        // val = ["type", null] → size 2, get(1) == null → return null
        FreeForm ff = new FreeForm();
        ff.put("Foo", createTypedValue(Arrays.asList("type", null)));

        String result = (String) extractFreeFormValue.invoke(resource, ff, "Foo");

        assertNull(result, "List where index 1 is null must return null");
    }

    @Test
    void extractFreeFormValue_valueIsListWithSizeOne_returnsNull() throws Exception {
        // val = ["only"] → size 1 < 2 → return null
        FreeForm ff = new FreeForm();
        ff.put("Foo", createTypedValue(List.of("only-one-element")));

        String result = (String) extractFreeFormValue.invoke(resource, ff, "Foo");

        assertNull(result,
                "List with size < 2 must return null (no value element at index 1)");
    }

    @Test
    void extractFreeFormValue_valueIsEmptyList_returnsNull() throws Exception {
        // val = [] → size 0 < 2 → return null
        FreeForm ff = new FreeForm();
        ff.put("Foo", createTypedValue(List.of()));

        String result = (String) extractFreeFormValue.invoke(resource, ff, "Foo");

        assertNull(result, "Empty list stored as value must return null (size < 2)");
    }
}
