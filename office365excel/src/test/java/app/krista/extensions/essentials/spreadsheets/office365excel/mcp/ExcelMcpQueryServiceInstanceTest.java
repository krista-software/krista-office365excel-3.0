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
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.lang.reflect.Method;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.lenient;

/**
 * Tests for instance methods of {@link ExcelMcpQueryService} that require
 * a constructed instance.
 *
 * <p>Covers:</p>
 * <ul>
 *   <li>{@code query()} — no AI key configured → proper FAILURE (no fallback)</li>
 *   <li>{@code resolvedPath} — /me/ → unchanged when no private auth configured</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
class ExcelMcpQueryServiceInstanceTest {

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

    // ── query() — no AI key ───────────────────────────────────────────────

    @Test
    void query_noAiKey_returnsFailure() {
        ExtensionResponse response = service.query("list my Excel files");
        assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
    }

    @Test
    void query_noAiKey_errorContainsAiKeyNotConfigured() {
        ExtensionResponse response = service.query("list my Excel files");
        assertNotNull(response.getError());
        assertTrue(String.valueOf(response.getError()).contains("AI API key is not configured"),
                "Error must say 'AI API key is not configured'");
    }

    @Test
    void query_noAiKey_errorMentionsAiAssistantTab() {
        ExtensionResponse response = service.query("list my Excel files");
        assertTrue(String.valueOf(response.getError()).contains("AI Assistant"),
                "Error must guide user to the AI Assistant tab to configure the key");
    }

    @Test
    void query_noAiKey_doesNotMentionAuthenticationConfiguration() {
        ExtensionResponse response = service.query("list my Excel files");
        assertFalse(String.valueOf(response.getError()).contains("authentication configuration"),
                "Missing AI key error must NOT mention 'authentication configuration'");
    }

    @Test
    void query_noAiKey_differentQuestionsAllReturnFailure() {
        // Regardless of the question, a missing AI key must always return FAILURE
        assertEquals(ExtensionResponse.Result.FAILURE,
                service.query("read Budget.xlsx").getResult());
        assertEquals(ExtensionResponse.Result.FAILURE,
                service.query("update cell A1 in Sales.xlsx").getResult());
        assertEquals(ExtensionResponse.Result.FAILURE,
                service.query("list all worksheets in Report.xlsm").getResult());
    }

    // ── resolvedPath — /me/ substitution ─────────────────────────────────

    @Test
    void resolvedPath_nullAttributes_returnsPathUnchanged() throws Exception {
        // attributeStore.load() returns null → no private auth configured
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "resolvedPath", String.class,
                app.krista.sdk.msauth.extension.mcp.DefaultMsAuthBridge.class);
        m.setAccessible(true);
        String result = (String) m.invoke(service, "/me/drive/root/children?$top=50", null);
        assertEquals("/me/drive/root/children?$top=50", result,
                "Without private auth attributes, path must remain unchanged");
    }

    @Test
    void resolvedPath_worksheetPath_unchangedWithoutPrivateAuth() throws Exception {
        Method m = ExcelMcpQueryService.class.getDeclaredMethod(
                "resolvedPath", String.class,
                app.krista.sdk.msauth.extension.mcp.DefaultMsAuthBridge.class);
        m.setAccessible(true);
        String path = "/me/drive/items/abc-123/workbook/worksheets";
        String result = (String) m.invoke(service, path, null);
        assertEquals(path, result,
                "Workbook worksheet path must remain unchanged without private auth");
    }
}
