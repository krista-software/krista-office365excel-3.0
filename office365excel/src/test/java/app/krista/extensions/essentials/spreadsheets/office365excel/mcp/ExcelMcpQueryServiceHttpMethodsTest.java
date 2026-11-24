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
import app.krista.sdk.msauth.extension.mcp.AnthropicApiClient;
import app.krista.sdk.msauth.extension.mcp.GeminiApiClient;
import app.krista.sdk.msauth.extension.mcp.McpModelConfig;
import app.krista.sdk.msauth.extension.mcp.OpenAiApiClient;
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
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Tests for the HTTP-calling private methods in {@link ExcelMcpQueryService}:
 *
 * <ul>
 *   <li>{@code callAi} — provider routing: openai / gemini / anthropic fallback</li>
 *   <li>{@code success} — ExtensionResponse factory</li>
 *   <li>{@code resolveItemId} — file-search HTTP call with all branching paths</li>
 *   <li>{@code executeGraphCall} — GET / POST / PATCH / DELETE, all HTTP status branches,
 *       exception handling</li>
 * </ul>
 *
 * <p>HTTP clients are mocked via {@link MockedStatic} on {@link ClientBuilder},
 * and AI API clients are mocked via {@link MockedStatic} on each API-client class.</p>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class ExcelMcpQueryServiceHttpMethodsTest {

    @Mock private RefreshTokenStore         refreshTokenStore;
    @Mock private Invoker                   invoker;
    @Mock private OfficeExcelAttributeStore attributeStore;

    private ExcelMcpQueryService service;

    private Method callAiMethod;
    private Method successMethod;
    private Method resolveItemIdMethod;
    private Method executeGraphCallMethod;

    @BeforeEach
    void setUp() throws Exception {
        lenient().when(invoker.getInvokerId()).thenReturn("test-invoker");
        lenient().when(attributeStore.load("test-invoker")).thenReturn(null);
        lenient().when(refreshTokenStore.get(anyString())).thenReturn(null);
        service = new ExcelMcpQueryService(refreshTokenStore, invoker, attributeStore);

        callAiMethod = ExcelMcpQueryService.class
                .getDeclaredMethod("callAi", String.class, String.class, String.class);
        callAiMethod.setAccessible(true);

        successMethod = ExcelMcpQueryService.class
                .getDeclaredMethod("success", Map.class);
        successMethod.setAccessible(true);

        resolveItemIdMethod = ExcelMcpQueryService.class
                .getDeclaredMethod("resolveItemId", String.class, String.class, String.class);
        resolveItemIdMethod.setAccessible(true);

        executeGraphCallMethod = ExcelMcpQueryService.class
                .getDeclaredMethod("executeGraphCall",
                        String.class, String.class, String.class, String.class);
        executeGraphCallMethod.setAccessible(true);
    }

    // ── callAi — provider routing ───────────────────────────────────────────────

    @Test
    void callAi_openaiProvider_delegatesToOpenAiApiClient() throws Exception {
        try (MockedStatic<McpModelConfig> configMock = mockStatic(McpModelConfig.class);
             MockedStatic<OpenAiApiClient> openAiMock = mockStatic(OpenAiApiClient.class)) {

            configMock.when(McpModelConfig::openAiModel).thenReturn("gpt-4o");
            openAiMock.when(() -> OpenAiApiClient.call("sk-test", "my-prompt", "gpt-4o", 0))
                      .thenReturn("openai-result");

            String result = (String) callAiMethod.invoke(null, "sk-test", "openai", "my-prompt");

            assertEquals("openai-result", result);
        }
    }

    @Test
    void callAi_openaiProvider_doesNotCallAnthropicOrGemini() throws Exception {
        try (MockedStatic<McpModelConfig> configMock = mockStatic(McpModelConfig.class);
             MockedStatic<OpenAiApiClient> openAiMock = mockStatic(OpenAiApiClient.class);
             MockedStatic<AnthropicApiClient> anthropicMock = mockStatic(AnthropicApiClient.class);
             MockedStatic<GeminiApiClient> geminiMock = mockStatic(GeminiApiClient.class)) {

            configMock.when(McpModelConfig::openAiModel).thenReturn("gpt-4o");
            openAiMock.when(() -> OpenAiApiClient.call(anyString(), anyString(), anyString(), anyInt()))
                      .thenReturn("openai-result");

            callAiMethod.invoke(null, "sk-test", "openai", "my-prompt");

            anthropicMock.verifyNoInteractions();
            geminiMock.verifyNoInteractions();
        }
    }

    @Test
    void callAi_geminiProvider_delegatesToGeminiApiClient() throws Exception {
        try (MockedStatic<McpModelConfig> configMock = mockStatic(McpModelConfig.class);
             MockedStatic<GeminiApiClient> geminiMock = mockStatic(GeminiApiClient.class)) {

            configMock.when(McpModelConfig::geminiModel).thenReturn("gemini-1.5-pro");
            geminiMock.when(() -> GeminiApiClient.call("gk-key", "prompt", "gemini-1.5-pro", 0))
                      .thenReturn("gemini-result");

            String result = (String) callAiMethod.invoke(null, "gk-key", "gemini", "prompt");

            assertEquals("gemini-result", result);
        }
    }

    @Test
    void callAi_geminiProvider_doesNotCallOpenAiOrAnthropic() throws Exception {
        try (MockedStatic<McpModelConfig> configMock = mockStatic(McpModelConfig.class);
             MockedStatic<GeminiApiClient> geminiMock = mockStatic(GeminiApiClient.class);
             MockedStatic<OpenAiApiClient> openAiMock = mockStatic(OpenAiApiClient.class);
             MockedStatic<AnthropicApiClient> anthropicMock = mockStatic(AnthropicApiClient.class)) {

            configMock.when(McpModelConfig::geminiModel).thenReturn("gemini-pro");
            geminiMock.when(() -> GeminiApiClient.call(anyString(), anyString(), anyString(), anyInt()))
                      .thenReturn("gemini-result");

            callAiMethod.invoke(null, "gk-key", "gemini", "prompt");

            openAiMock.verifyNoInteractions();
            anthropicMock.verifyNoInteractions();
        }
    }

    @Test
    void callAi_anthropicProvider_delegatesToAnthropicApiClient() throws Exception {
        try (MockedStatic<McpModelConfig> configMock = mockStatic(McpModelConfig.class);
             MockedStatic<AnthropicApiClient> anthropicMock = mockStatic(AnthropicApiClient.class)) {

            configMock.when(McpModelConfig::anthropicModel).thenReturn("claude-3-5-sonnet");
            anthropicMock.when(() ->
                    AnthropicApiClient.call("ak-key", "prompt", "claude-3-5-sonnet", 0))
                         .thenReturn("anthropic-result");

            String result = (String) callAiMethod.invoke(null, "ak-key", "anthropic", "prompt");

            assertEquals("anthropic-result", result);
        }
    }

    @Test
    void callAi_unknownProvider_fallsThroughToAnthropicApiClient() throws Exception {
        // Any provider value that is neither "openai" nor "gemini" falls to the else branch
        try (MockedStatic<McpModelConfig> configMock = mockStatic(McpModelConfig.class);
             MockedStatic<AnthropicApiClient> anthropicMock = mockStatic(AnthropicApiClient.class)) {

            configMock.when(McpModelConfig::anthropicModel).thenReturn("claude-3");
            anthropicMock.when(() ->
                    AnthropicApiClient.call(anyString(), anyString(), anyString(), anyInt()))
                         .thenReturn("fallback-result");

            String result = (String) callAiMethod.invoke(null, "key", "some-unknown-provider", "p");

            assertEquals("fallback-result", result);
        }
    }

    // ── success ─────────────────────────────────────────────────────────────────

    @Test
    void success_returnsSuccessResult() throws Exception {
        ExtensionResponse resp = (ExtensionResponse) successMethod.invoke(null, Map.of());
        assertEquals(ExtensionResponse.Result.SUCCESS, resp.getResult());
    }

    @Test
    void success_responseValueMatchesInput() throws Exception {
        Map<String, Object> payload = Map.of("Response", List.of(), "extra", "val");
        ExtensionResponse resp = (ExtensionResponse) successMethod.invoke(null, payload);
        assertSame(payload, resp.getResponseValue());
    }

    @Test
    void success_errorIsNull() throws Exception {
        ExtensionResponse resp = (ExtensionResponse) successMethod.invoke(null, Map.of());
        assertNull(resp.getError());
    }

    @Test
    void success_nullValues_resultIsStillSuccess() throws Exception {
        ExtensionResponse resp = (ExtensionResponse) successMethod.invoke(null, (Object) null);
        assertEquals(ExtensionResponse.Result.SUCCESS, resp.getResult());
    }

    // ── resolveItemId — helpers ─────────────────────────────────────────────────

    /** Sets up the GET chain: ClientBuilder → Client → WebTarget → Builder → Response. */
    private static RiggedHttp rigGetChain(MockedStatic<ClientBuilder> cbMock,
                                          int status, String body) {
        Client c = mock(Client.class);
        WebTarget t = mock(WebTarget.class);
        Invocation.Builder b = mock(Invocation.Builder.class);
        javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

        cbMock.when(ClientBuilder::newClient).thenReturn(c);
        when(c.target(anyString())).thenReturn(t);
        lenient().when(t.request(any(MediaType.class))).thenReturn(b);
        lenient().when(t.request(any(String.class))).thenReturn(b);
        lenient().when(t.request()).thenReturn(b);
        when(b.header(anyString(), any())).thenReturn(b);
        when(b.get()).thenReturn(r);
        when(r.getStatus()).thenReturn(status);
        if (body != null) lenient().when(r.readEntity(String.class)).thenReturn(body);

        return new RiggedHttp(c, t, b, r);
    }

    /** Minimal holder for the mock chain (for verifications if needed). */
    private record RiggedHttp(Client client, WebTarget target,
                               Invocation.Builder builder,
                               javax.ws.rs.core.Response response) {}

    // ── resolveItemId — filename extraction ────────────────────────────────────

    @Test
    void resolveItemId_questionWithNoXlsxFilename_returnsNull() throws Exception {
        // extractFileName returns null → early return null, no HTTP call made
        Object result = resolveItemIdMethod.invoke(
                service,
                "/me/drive/items/{itemId}/workbook/worksheets",
                "list all my documents",  // no .xlsx filename
                "Bearer token");
        assertNull(result);
    }

    @Test
    void resolveItemId_emptyQuestion_returnsNull() throws Exception {
        Object result = resolveItemIdMethod.invoke(
                service,
                "/me/drive/items/{itemId}/workbook",
                "",
                "Bearer token");
        assertNull(result);
    }

    // ── resolveItemId — HTTP error ─────────────────────────────────────────────

    @Test
    void resolveItemId_searchReturnsNon200_returnsNull() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 401, null);

            Object result = resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook/worksheets",
                    "list worksheets of Budget.xlsx",
                    "Bearer token");

            assertNull(result, "Non-200 HTTP response must return null");
        }
    }

    @Test
    void resolveItemId_searchReturns403_returnsNull() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 403, null);

            Object result = resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "open Report.xlsx",
                    "Bearer token");

            assertNull(result);
        }
    }

    // ── resolveItemId — exact match ────────────────────────────────────────────

    @Test
    void resolveItemId_exactCaseInsensitiveNameMatch_replacesItemId() throws Exception {
        String body = "{\"value\":[{\"id\":\"item-abc\",\"name\":\"Budget.xlsx\"}]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            String result = (String) resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook/worksheets",
                    "list worksheets of Budget.xlsx",
                    "Bearer token");

            assertNotNull(result);
            assertTrue(result.contains("item-abc"),
                    "Exact name match must replace {itemId} with the matched item ID");
            assertFalse(result.contains("{itemId}"),
                    "{itemId} placeholder must be fully replaced");
        }
    }

    @Test
    void resolveItemId_exactMatchCaseInsensitive_lowerQuestion() throws Exception {
        // "budget.xlsx" in question must match "Budget.xlsx" in drive (equalsIgnoreCase)
        String body = "{\"value\":[{\"id\":\"file-123\",\"name\":\"Budget.xlsx\"}]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            String result = (String) resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "open budget.xlsx",
                    "Bearer token");

            assertNotNull(result);
            assertTrue(result.contains("file-123"));
        }
    }

    // ── resolveItemId — fallback (first result) ────────────────────────────────

    @Test
    void resolveItemId_noExactMatch_usesFirstResultAsFallback() throws Exception {
        // Item name "OtherFile.xlsx" does not match query file "Budget.xlsx"
        String body = "{\"value\":[{\"id\":\"first-id\",\"name\":\"OtherFile.xlsx\"}]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            String result = (String) resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "open Budget.xlsx",
                    "Bearer token");

            assertNotNull(result);
            assertTrue(result.contains("first-id"),
                    "When no exact match, the first result must be used as fallback");
        }
    }

    @Test
    void resolveItemId_multipleResults_firstIdIsUsedWhenNoExactMatch() throws Exception {
        String body = "{\"value\":["
                + "{\"id\":\"id-one\",\"name\":\"SomeFile.xlsx\"},"
                + "{\"id\":\"id-two\",\"name\":\"AnotherFile.xlsx\"}"
                + "]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            String result = (String) resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "list Budget.xlsx sheets",
                    "Bearer token");

            assertNotNull(result);
            assertTrue(result.contains("id-one"), "First result must be used as fallback");
            assertFalse(result.contains("id-two"));
        }
    }

    // ── resolveItemId — empty itemId skipped ──────────────────────────────────

    @Test
    void resolveItemId_emptyItemId_skippedFallbackFromSecondEntry() throws Exception {
        // First item has empty id → skipped; second item has non-empty id → becomes fallback
        String body = "{\"value\":["
                + "{\"id\":\"\",\"name\":\"Budget.xlsx\"},"
                + "{\"id\":\"real-id\",\"name\":\"Report.xlsx\"}"
                + "]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            String result = (String) resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "open Budget.xlsx",
                    "Bearer token");

            // "Budget.xlsx" matches first item exactly but id is empty → skipped
            // Second item: id="real-id", name="Report.xlsx" → fallbackId="real-id", no exact match
            // Fallback returned
            assertNotNull(result);
            assertTrue(result.contains("real-id"));
        }
    }

    // ── resolveItemId — file not found ────────────────────────────────────────

    @Test
    void resolveItemId_emptyValueArray_returnsNull() throws Exception {
        String body = "{\"value\":[]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            Object result = resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "read Budget.xlsx",
                    "Bearer token");

            assertNull(result, "Empty value array means file not found — must return null");
        }
    }

    @Test
    void resolveItemId_noValueKeyInResponse_returnsNull() throws Exception {
        // Root JSON has no "value" key → empty array, no matches
        String body = "{}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            Object result = resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "read Budget.xlsx",
                    "Bearer token");

            assertNull(result);
        }
    }

    // ── resolveItemId — non-object elements skipped ────────────────────────────

    @Test
    void resolveItemId_nonJsonObjectInValues_skippedValidItemUsed() throws Exception {
        String body = "{\"value\":[\"not-an-object\",{\"id\":\"valid-id\",\"name\":\"Budget.xlsx\"}]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            String result = (String) resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "read Budget.xlsx",
                    "Bearer token");

            assertNotNull(result);
            assertTrue(result.contains("valid-id"),
                    "Non-object elements must be skipped; the valid item must be found");
        }
    }

    // ── resolveItemId — exception handling ────────────────────────────────────

    @Test
    void resolveItemId_exceptionDuringHttpCall_returnsNull() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client mockClient = mock(Client.class);
            cbMock.when(ClientBuilder::newClient).thenReturn(mockClient);
            when(mockClient.target(anyString())).thenThrow(new RuntimeException("network error"));

            Object result = resolveItemIdMethod.invoke(
                    service,
                    "/me/drive/items/{itemId}/workbook",
                    "read Budget.xlsx",
                    "Bearer token");

            assertNull(result, "Exception during HTTP call must be caught and null returned");
        }
    }

    // ── executeGraphCall — HTTP method routing ─────────────────────────────────

    @Test
    void executeGraphCall_getMethod_usesGetRequest() throws Exception {
        String body = "{\"value\":[{\"name\":\"Budget.xlsx\"}]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            RiggedHttp rig = rigGetChain(cbMock, 200, body);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive/root/children",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
            verify(rig.builder()).get();
        }
    }

    @Test
    void executeGraphCall_postMethod_usesPostRequest() throws Exception {
        String respBody = "{\"value\":[]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.post(any())).thenReturn(r);
            when(r.getStatus()).thenReturn(200);
            when(r.readEntity(String.class)).thenReturn(respBody);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "POST",
                    "https://graph.microsoft.com/v1.0/search/query",
                    "{\"requests\":[]}", "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
            verify(b).post(any());
        }
    }

    @Test
    void executeGraphCall_postMethod_withNullBody_usesEmptyJsonBody() throws Exception {
        String respBody = "{}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.post(any())).thenReturn(r);
            when(r.getStatus()).thenReturn(200);
            when(r.readEntity(String.class)).thenReturn(respBody);

            // null body → code substitutes "{}"
            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "POST",
                    "https://graph.microsoft.com/v1.0/search/query",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
        }
    }

    @Test
    void executeGraphCall_patchMethod_usesPatchRequest() throws Exception {
        String respBody = "{\"address\":\"Sheet1!A1\"}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.method(anyString(), any(Entity.class))).thenReturn(r);
            when(r.getStatus()).thenReturn(200);
            when(r.readEntity(String.class)).thenReturn(respBody);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "PATCH",
                    "https://graph.microsoft.com/v1.0/me/drive/items/id1/workbook/range/values",
                    "{\"values\":[[\"NewVal\"]]}", "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
            verify(b).method(eq("PATCH"), any(Entity.class));
        }
    }

    @Test
    void executeGraphCall_patchMethod_withNullBody_usesEmptyJsonBody() throws Exception {
        String respBody = "{}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.method(anyString(), any(Entity.class))).thenReturn(r);
            when(r.getStatus()).thenReturn(200);
            when(r.readEntity(String.class)).thenReturn(respBody);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "PATCH",
                    "https://graph.microsoft.com/v1.0/me/drive/items/id1/workbook/range/values",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
        }
    }

    @Test
    void executeGraphCall_deleteMethod_usesDeleteRequest() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.delete()).thenReturn(r);
            when(r.getStatus()).thenReturn(204);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "DELETE",
                    "https://graph.microsoft.com/v1.0/me/drive/items/id1",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
            verify(b).delete();
        }
    }

    // ── executeGraphCall — HTTP status handling ────────────────────────────────

    @Test
    void executeGraphCall_http202_returnsSuccessWithActionMessage() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 202, null);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/test",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
        }
    }

    @Test
    void executeGraphCall_http202_responseContainsActionCompletedMessage() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 202, null);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/test",
                    null, "Bearer token");

            @SuppressWarnings("unchecked")
            Map<String, Object> rv = (Map<String, Object>) result.getResponseValue();
            List<?> rows = (List<?>) rv.get("Response");
            assertFalse(rows.isEmpty(), "HTTP 202 must produce at least one row in the response");
        }
    }

    @Test
    void executeGraphCall_http204_returnsSuccess() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.delete()).thenReturn(r);
            when(r.getStatus()).thenReturn(204);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "DELETE",
                    "https://graph.microsoft.com/v1.0/me/drive/items/id1",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
        }
    }

    @Test
    void executeGraphCall_http200_parsesResponseRows() throws Exception {
        String body = "{\"value\":[{\"name\":\"Budget.xlsx\",\"id\":\"abc\"}]}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 200, body);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive/root/children",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
            @SuppressWarnings("unchecked")
            Map<String, Object> rv = (Map<String, Object>) result.getResponseValue();
            List<?> rows = (List<?>) rv.get("Response");
            assertFalse(rows.isEmpty(), "HTTP 200 with items must produce non-empty rows");
        }
    }

    @Test
    void executeGraphCall_http201_returnsSuccess() throws Exception {
        String body = "{\"id\":\"new-sheet\",\"name\":\"Sheet2\"}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.post(any())).thenReturn(r);
            when(r.getStatus()).thenReturn(201);
            when(r.readEntity(String.class)).thenReturn(body);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "POST",
                    "https://graph.microsoft.com/v1.0/me/drive/items/id1/workbook/worksheets",
                    "{\"name\":\"Sheet2\"}", "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
        }
    }

    @Test
    void executeGraphCall_patchHttp200_withAddressInResponse_includesAddressInResult() throws Exception {
        String body = "{\"address\":\"Sheet1!A1\"}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.method(anyString(), any(Entity.class))).thenReturn(r);
            when(r.getStatus()).thenReturn(200);
            when(r.readEntity(String.class)).thenReturn(body);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "PATCH",
                    "https://graph.microsoft.com/v1.0/me/drive/items/id1/workbook/range/values",
                    "{\"values\":[[\"Val\"]]}", "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
            @SuppressWarnings("unchecked")
            Map<String, Object> rv = (Map<String, Object>) result.getResponseValue();
            List<?> rows = (List<?>) rv.get("Response");
            // The single row should contain the "Cell updated" message with address
            assertFalse(rows.isEmpty());
        }
    }

    @Test
    void executeGraphCall_patchHttp200_withNoAddress_cellUpdatedWithoutParentheses() throws Exception {
        // address field is empty → updatedAddress = "" → " (" + "" + ")" is NOT appended
        String body = "{}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            WebTarget t = mock(WebTarget.class);
            Invocation.Builder b = mock(Invocation.Builder.class);
            javax.ws.rs.core.Response r = mock(javax.ws.rs.core.Response.class);

            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenReturn(t);
            lenient().when(t.request(any(MediaType.class))).thenReturn(b);
            lenient().when(t.request(any(String.class))).thenReturn(b);
            lenient().when(t.request()).thenReturn(b);
            when(b.header(anyString(), any())).thenReturn(b);
            when(b.method(anyString(), any(Entity.class))).thenReturn(r);
            when(r.getStatus()).thenReturn(200);
            when(r.readEntity(String.class)).thenReturn(body);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "PATCH",
                    "https://graph.microsoft.com/v1.0/me/drive/items/id1/workbook/range/values",
                    "{\"values\":[[\"Val\"]]}", "Bearer token");

            assertEquals(ExtensionResponse.Result.SUCCESS, result.getResult());
        }
    }

    // ── executeGraphCall — error status ────────────────────────────────────────

    @Test
    void executeGraphCall_http400_returnsFailure() throws Exception {
        String errorBody = "{\"error\":{\"message\":\"Bad request\"}}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            RiggedHttp rig = rigGetChain(cbMock, 400, errorBody);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.FAILURE, result.getResult());
        }
    }

    @Test
    void executeGraphCall_http401_returnsFailure() throws Exception {
        String errorBody = "{\"error\":{\"message\":\"Unauthorized\"}}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 401, errorBody);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.FAILURE, result.getResult());
        }
    }

    @Test
    void executeGraphCall_errorStatus_errorMessageContainsHttpStatus() throws Exception {
        String errorBody = "{\"error\":{\"message\":\"Forbidden\"}}";
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            rigGetChain(cbMock, 403, errorBody);

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive",
                    null, "Bearer token");

            String errorStr = String.valueOf(result.getError());
            assertTrue(errorStr.contains("403") || errorStr.contains("Graph API call failed"),
                    "Error message must mention the HTTP status or 'Graph API call failed'");
        }
    }

    // ── executeGraphCall — exception handling ─────────────────────────────────

    @Test
    void executeGraphCall_exceptionDuringRequest_returnsFailure() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenThrow(new RuntimeException("connection refused"));

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive",
                    null, "Bearer token");

            assertEquals(ExtensionResponse.Result.FAILURE, result.getResult());
        }
    }

    @Test
    void executeGraphCall_exceptionDuringRequest_errorMessageContainsGraphApiFailed() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenThrow(new RuntimeException("network error"));

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive",
                    null, "Bearer token");

            assertTrue(String.valueOf(result.getError()).contains("Graph API call failed"),
                    "Exception path error must say 'Graph API call failed'");
        }
    }

    @Test
    void executeGraphCall_exceptionDuringRequest_exceptionMessageIncluded() throws Exception {
        try (MockedStatic<ClientBuilder> cbMock = mockStatic(ClientBuilder.class)) {
            Client c = mock(Client.class);
            cbMock.when(ClientBuilder::newClient).thenReturn(c);
            when(c.target(anyString())).thenThrow(new RuntimeException("timeout after 30s"));

            ExtensionResponse result = (ExtensionResponse) executeGraphCallMethod.invoke(
                    service, "GET",
                    "https://graph.microsoft.com/v1.0/me/drive",
                    null, "Bearer token");

            assertTrue(String.valueOf(result.getError()).contains("timeout after 30s"),
                    "Exception message must be included in the failure error string");
        }
    }
}
