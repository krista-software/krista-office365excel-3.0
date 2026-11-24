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
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.MockedStatic;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import javax.ws.rs.client.Client;
import javax.ws.rs.client.ClientBuilder;
import javax.ws.rs.client.Entity;
import javax.ws.rs.client.Invocation;
import javax.ws.rs.client.WebTarget;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import java.lang.reflect.Method;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.*;

/**
 * Tests for the private HTTP helper methods ({@code graphPost}, {@code graphGet})
 * and the tool execution methods ({@code execListWorkbooks}, {@code execSearchWorkbooks},
 * {@code execGetWorkbook}, {@code execListWorksheets}, {@code execReadWorksheet})
 * of {@link ExcelMcpProtocolResource}.
 *
 * <p>Uses {@code MockedStatic<ClientBuilder>} to intercept JAX-RS client creation
 * so no real HTTP calls are made.</p>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class ExcelMcpProtocolResourceHttpTest {

    @Mock private RefreshTokenStore         refreshTokenStore;
    @Mock private Invoker                   invoker;
    @Mock private OfficeExcelAttributeStore attributeStore;

    @Mock private Client          client;
    @Mock private WebTarget       webTarget;
    @Mock private Invocation.Builder builder;

    private MockedStatic<ClientBuilder> clientBuilderStatic;
    private ExcelMcpProtocolResource resource;

    @BeforeEach
    void setUp() throws Exception {
        lenient().when(invoker.getInvokerId()).thenReturn("test-invoker");
        lenient().when(attributeStore.load("test-invoker")).thenReturn(null);
        lenient().when(refreshTokenStore.get(anyString())).thenReturn(null);

        resource = new ExcelMcpProtocolResource(refreshTokenStore, invoker, attributeStore);

        // Mock the JAX-RS client chain
        clientBuilderStatic = mockStatic(ClientBuilder.class);
        clientBuilderStatic.when(ClientBuilder::newClient).thenReturn(client);
        lenient().when(client.target(anyString())).thenReturn(webTarget);
        lenient().when(webTarget.request(MediaType.APPLICATION_JSON)).thenReturn(builder);
        lenient().when(builder.header(anyString(), anyString())).thenReturn(builder);
    }

    @AfterEach
    void tearDown() {
        clientBuilderStatic.close();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // graphPost tests
    // ══════════════════════════════════════════════════════════════════════════

    @Test
    void graphPost_status200_returnsBody() throws Exception {
        Response resp = mockResponse(200, "{\"ok\":true}");
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeGraphPost("https://graph.microsoft.com/v1.0/test", "{}", "Bearer token");

        assertEquals("{\"ok\":true}", result);
        verify(client).close();
    }

    @Test
    void graphPost_status201_returnsBody() throws Exception {
        Response resp = mockResponse(201, "{\"created\":true}");
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeGraphPost("https://graph.microsoft.com/v1.0/test", "{}", "Bearer token");

        assertEquals("{\"created\":true}", result);
    }

    @Test
    void graphPost_status400_returnsWrappedError() throws Exception {
        Response resp = mockResponse(400, "Bad Request body");
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeGraphPost("https://graph.microsoft.com/v1.0/test", "{}", "Bearer token");

        assertNotNull(result);
        assertTrue(result.contains("error"), "Non-200/201 should return wrapped error");
    }

    @Test
    void graphPost_status403_withGraphError_passesThrough() throws Exception {
        String graphErr = "{\"error\":{\"code\":\"AccessDenied\",\"message\":\"Forbidden\"}}";
        Response resp = mockResponse(403, graphErr);
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeGraphPost("https://graph.microsoft.com/v1.0/test", "{}", "Bearer token");

        assertTrue(result.contains("AccessDenied"));
    }

    @Test
    void graphPost_status500_returnsWrappedErrorWithStatusCode() throws Exception {
        Response resp = mockResponse(500, "Internal Server Error");
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeGraphPost("https://graph.microsoft.com/v1.0/test", "{}", "Bearer token");

        assertTrue(result.contains("500"));
    }

    @Test
    void graphPost_exception_returnsNull() throws Exception {
        when(builder.post(any(Entity.class))).thenThrow(new RuntimeException("connection refused"));

        String result = invokeGraphPost("https://graph.microsoft.com/v1.0/test", "{}", "Bearer token");

        assertNull(result);
        verify(client).close();
    }

    @Test
    void graphPost_alwaysClosesClient() throws Exception {
        Response resp = mockResponse(200, "ok");
        when(builder.post(any(Entity.class))).thenReturn(resp);

        invokeGraphPost("https://graph.microsoft.com/v1.0/test", "{}", "Bearer token");

        verify(client).close();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // graphGet tests
    // ══════════════════════════════════════════════════════════════════════════

    @Test
    void graphGet_status200_returnsBody() throws Exception {
        Response resp = mockResponse(200, "{\"value\":[]}");
        when(builder.get()).thenReturn(resp);

        String result = invokeGraphGet("https://graph.microsoft.com/v1.0/me/drive", "Bearer token");

        assertEquals("{\"value\":[]}", result);
        verify(client).close();
    }

    @Test
    void graphGet_status404_returnsWrappedError() throws Exception {
        Response resp = mockResponse(404, "Not Found");
        when(builder.get()).thenReturn(resp);

        String result = invokeGraphGet("https://graph.microsoft.com/v1.0/me/drive", "Bearer token");

        assertNotNull(result);
        assertTrue(result.contains("error"));
    }

    @Test
    void graphGet_status401_withGraphError_passesThrough() throws Exception {
        String graphErr = "{\"error\":{\"code\":\"Unauthorized\",\"message\":\"Token expired\"}}";
        Response resp = mockResponse(401, graphErr);
        when(builder.get()).thenReturn(resp);

        String result = invokeGraphGet("https://graph.microsoft.com/v1.0/me/drive", "Bearer token");

        assertTrue(result.contains("Token expired"));
    }

    @Test
    void graphGet_exception_returnsNull() throws Exception {
        when(builder.get()).thenThrow(new RuntimeException("timeout"));

        String result = invokeGraphGet("https://graph.microsoft.com/v1.0/me/drive", "Bearer token");

        assertNull(result);
        verify(client).close();
    }

    @Test
    void graphGet_alwaysClosesClient() throws Exception {
        Response resp = mockResponse(200, "ok");
        when(builder.get()).thenReturn(resp);

        invokeGraphGet("https://graph.microsoft.com/v1.0/me/drive", "Bearer token");

        verify(client).close();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // execListWorkbooks tests
    // ══════════════════════════════════════════════════════════════════════════

    @Test
    void execListWorkbooks_successfulResponse_formatsWorkbooks() throws Exception {
        String searchResponse = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":\"id1\","
                + "\"resource\":{\"name\":\"Budget.xlsx\",\"size\":\"2048\"}"
                + "}]}]}]}";
        Response resp = mockResponse(200, searchResponse);
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeExecListWorkbooks("Bearer token");

        assertTrue(result.contains("Budget.xlsx"));
    }

    @Test
    void execListWorkbooks_nullResponse_returnsError() throws Exception {
        when(builder.post(any(Entity.class))).thenThrow(new RuntimeException("fail"));

        String result = invokeExecListWorkbooks("Bearer token");

        assertEquals("Error: failed to reach Microsoft Graph API.", result);
    }

    @Test
    void execListWorkbooks_emptySearchResults_returnsNoWorkbooks() throws Exception {
        Response resp = mockResponse(200, "{\"value\":[{\"hitsContainers\":[{\"hits\":[]}]}]}");
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeExecListWorkbooks("Bearer token");

        assertEquals("No workbooks found.", result);
    }

    // ══════════════════════════════════════════════════════════════════════════
    // execSearchWorkbooks tests
    // ══════════════════════════════════════════════════════════════════════════

    @Test
    void execSearchWorkbooks_successfulResponse_formatsResults() throws Exception {
        String searchResponse = "{\"value\":[{\"hitsContainers\":[{\"hits\":[{"
                + "\"hitId\":\"id2\","
                + "\"resource\":{\"name\":\"Sales.xlsx\"}"
                + "}]}]}]}";
        Response resp = mockResponse(200, searchResponse);
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeExecSearchWorkbooks("Bearer token", "sales");

        assertTrue(result.contains("Sales.xlsx"));
        assertTrue(result.contains("workbooks matching \"sales\""));
    }

    @Test
    void execSearchWorkbooks_nullResponse_returnsError() throws Exception {
        when(builder.post(any(Entity.class))).thenThrow(new RuntimeException("fail"));

        String result = invokeExecSearchWorkbooks("Bearer token", "test");

        assertEquals("Error: failed to reach Microsoft Graph API.", result);
    }

    @Test
    void execSearchWorkbooks_noResults_returnsNoMatchMessage() throws Exception {
        Response resp = mockResponse(200, "{\"value\":[]}");
        when(builder.post(any(Entity.class))).thenReturn(resp);

        String result = invokeExecSearchWorkbooks("Bearer token", "nonexistent");

        assertTrue(result.contains("No workbooks matching \"nonexistent\" found."));
    }

    // ══════════════════════════════════════════════════════════════════════════
    // execGetWorkbook tests
    // ══════════════════════════════════════════════════════════════════════════

    @Test
    void execGetWorkbook_successfulResponse_formatsSingleItem() throws Exception {
        String itemJson = "{\"name\":\"Report.xlsx\",\"id\":\"file-123\",\"size\":\"1048576\"}";
        Response resp = mockResponse(200, itemJson);
        when(builder.get()).thenReturn(resp);

        String result = invokeExecGetWorkbook("Bearer token", "file-123");

        assertTrue(result.contains("Report.xlsx"));
        assertTrue(result.contains("file-123"));
    }

    @Test
    void execGetWorkbook_nullResponse_returnsError() throws Exception {
        when(builder.get()).thenThrow(new RuntimeException("fail"));

        String result = invokeExecGetWorkbook("Bearer token", "file-123");

        assertEquals("Error: failed to reach Microsoft Graph API.", result);
    }

    @Test
    void execGetWorkbook_graphError_returnsErrorMessage() throws Exception {
        String errorJson = "{\"error\":{\"message\":\"Item not found\"}}";
        Response resp = mockResponse(200, errorJson);
        when(builder.get()).thenReturn(resp);

        String result = invokeExecGetWorkbook("Bearer token", "bad-id");

        assertTrue(result.contains("Item not found"));
    }

    @Test
    void execGetWorkbook_usesResolvedPath_defaultNoAttrs() throws Exception {
        Response resp = mockResponse(200, "{\"name\":\"F.xlsx\",\"id\":\"x\"}");
        when(builder.get()).thenReturn(resp);

        invokeExecGetWorkbook("Bearer token", "abc");

        // With null attrs, path should stay /me/...
        verify(client).target("https://graph.microsoft.com/v1.0/me/drive/items/abc");
    }

    @Test
    void execGetWorkbook_usesResolvedPath_privateAuth() throws Exception {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn("PRIVATE");
        when(attrs.getEmail()).thenReturn("user@example.com");
        when(attributeStore.load("test-invoker")).thenReturn(attrs);

        Response resp = mockResponse(200, "{\"name\":\"F.xlsx\",\"id\":\"x\"}");
        when(builder.get()).thenReturn(resp);

        invokeExecGetWorkbook("Bearer token", "abc");

        verify(client).target("https://graph.microsoft.com/v1.0/users/user@example.com/drive/items/abc");
    }

    // ══════════════════════════════════════════════════════════════════════════
    // execListWorksheets tests
    // ══════════════════════════════════════════════════════════════════════════

    @Test
    void execListWorksheets_successfulResponse_formatsWorksheets() throws Exception {
        String json = "{\"value\":[{\"name\":\"Sheet1\",\"visibility\":\"Visible\"}]}";
        Response resp = mockResponse(200, json);
        when(builder.get()).thenReturn(resp);

        String result = invokeExecListWorksheets("Bearer token", "file-1");

        assertTrue(result.contains("Sheet1"));
        assertTrue(result.contains("1 worksheet(s)"));
    }

    @Test
    void execListWorksheets_nullResponse_returnsError() throws Exception {
        when(builder.get()).thenThrow(new RuntimeException("fail"));

        String result = invokeExecListWorksheets("Bearer token", "file-1");

        assertEquals("Error: failed to reach Microsoft Graph API.", result);
    }

    @Test
    void execListWorksheets_emptyValue_returnsNoWorksheetsMessage() throws Exception {
        Response resp = mockResponse(200, "{\"value\":[]}");
        when(builder.get()).thenReturn(resp);

        String result = invokeExecListWorksheets("Bearer token", "file-1");

        assertEquals("No worksheets found in this workbook.", result);
    }

    // ══════════════════════════════════════════════════════════════════════════
    // execReadWorksheet tests
    // ══════════════════════════════════════════════════════════════════════════

    @Test
    void execReadWorksheet_successfulResponse_formatsData() throws Exception {
        String json = "{\"values\":[[\"Name\",\"Age\"],[\"Alice\",\"30\"]]}";
        Response resp = mockResponse(200, json);
        when(builder.get()).thenReturn(resp);

        String result = invokeExecReadWorksheet("Bearer token", "file-1", "Sheet1");

        assertTrue(result.contains("Name"));
        assertTrue(result.contains("Alice"));
        assertTrue(result.contains("Sheet1"));
    }

    @Test
    void execReadWorksheet_nullResponse_returnsError() throws Exception {
        when(builder.get()).thenThrow(new RuntimeException("fail"));

        String result = invokeExecReadWorksheet("Bearer token", "file-1", "Sheet1");

        assertEquals("Error: failed to reach Microsoft Graph API.", result);
    }

    @Test
    void execReadWorksheet_emptyValues_returnsEmptyMessage() throws Exception {
        Response resp = mockResponse(200, "{\"values\":[]}");
        when(builder.get()).thenReturn(resp);

        String result = invokeExecReadWorksheet("Bearer token", "file-1", "Data");

        assertEquals("Sheet \"Data\" is empty.", result);
    }

    @Test
    void execReadWorksheet_sheetNameWithSpaces_encodesSpaces() throws Exception {
        Response resp = mockResponse(200, "{\"values\":[[\"Col\"],[\"Val\"]]}");
        when(builder.get()).thenReturn(resp);

        invokeExecReadWorksheet("Bearer token", "file-1", "My Sheet");

        // Verify the URL has %20 for spaces
        verify(client).target(contains("My%20Sheet"));
    }

    @Test
    void execReadWorksheet_usesResolvedPath() throws Exception {
        Response resp = mockResponse(200, "{\"values\":[[\"A\"],[\"B\"]]}");
        when(builder.get()).thenReturn(resp);

        invokeExecReadWorksheet("Bearer token", "xyz", "Sheet1");

        verify(client).target(
                "https://graph.microsoft.com/v1.0/me/drive/items/xyz/workbook/worksheets/Sheet1/usedRange");
    }

    // ══════════════════════════════════════════════════════════════════════════
    // Private helpers
    // ══════════════════════════════════════════════════════════════════════════

    private static Response mockResponse(int status, String body) {
        Response resp = mock(Response.class);
        when(resp.getStatus()).thenReturn(status);
        when(resp.readEntity(String.class)).thenReturn(body);
        return resp;
    }

    private static String contains(String substring) {
        return argThat(arg -> arg != null && arg.contains(substring));
    }

    // ── Reflection invokers ─────────────────────────────────────────────────

    private static String invokeGraphPost(String url, String body, String authHeader) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "graphPost", String.class, String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, url, body, authHeader);
    }

    private static String invokeGraphGet(String url, String authHeader) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "graphGet", String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(null, url, authHeader);
    }

    private String invokeExecListWorkbooks(String authHeader) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "execListWorkbooks", String.class);
        m.setAccessible(true);
        return (String) m.invoke(resource, authHeader);
    }

    private String invokeExecSearchWorkbooks(String authHeader, String keyword) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "execSearchWorkbooks", String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(resource, authHeader, keyword);
    }

    private String invokeExecGetWorkbook(String authHeader, String fileId) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "execGetWorkbook", String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(resource, authHeader, fileId);
    }

    private String invokeExecListWorksheets(String authHeader, String fileId) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "execListWorksheets", String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(resource, authHeader, fileId);
    }

    private String invokeExecReadWorksheet(String authHeader, String fileId, String sheet) throws Exception {
        Method m = ExcelMcpProtocolResource.class.getDeclaredMethod(
                "execReadWorksheet", String.class, String.class, String.class);
        m.setAccessible(true);
        return (String) m.invoke(resource, authHeader, fileId, sheet);
    }
}
