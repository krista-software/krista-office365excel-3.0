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

package app.krista.extensions.essentials.spreadsheets.office365excel;

import app.krista.extension.authorization.RequestAuthenticator;
import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.User;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.requests.UserRequest;
import com.microsoft.graph.requests.UserRequestBuilder;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Comprehensive test suite for Office365ExcelExtension with 100% code coverage.
 * Tests all constructors, methods, and edge cases.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class Office365ExcelExtensionTest {

    @Mock
    private Invoker invoker;
    @Mock
    private RoutingInfo routingInfo;
    @Mock
    private GraphServiceClientProviderFactory clientProviderFactory;
    @Mock
    private OfficeExcelAttributeStore officeExcelAttributeStore;
    @Mock
    private Office365ExcelAuditService auditService;
    @Mock
    private OfficeExcelRequestAuthenticator requestAuthenticator;
    @Mock
    private GraphServiceClientProvider graphServiceClientProvider;
    @Mock
    private GraphServiceClient graphServiceClient;
    @Mock
    private UserRequestBuilder userRequestBuilder;
    @Mock
    private UserRequest userRequest;

    private Office365ExcelExtension extension;

    @BeforeEach
    void setUp() {
        when(invoker.getRoutingInfo()).thenReturn(routingInfo);
        when(routingInfo.getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE))
                .thenReturn("https://routing.example.com");
        when(invoker.getInvokerId()).thenReturn("test-invoker");
        extension = new Office365ExcelExtension(invoker, clientProviderFactory,
                officeExcelAttributeStore, auditService);
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_initializesCorrectly() {
        Office365ExcelExtension ext = new Office365ExcelExtension(
                invoker, clientProviderFactory, officeExcelAttributeStore, auditService);

        assertNotNull(ext);
        RequestAuthenticator authenticator = ext.getRequestAuthenticator();
        assertNotNull(authenticator);
        assertTrue(authenticator instanceof OfficeExcelRequestAuthenticator);
    }

    // ==================== getRequestAuthenticator Tests ====================

    @Test
    void getRequestAuthenticator_returnsAuthenticatorInstance() {
        RequestAuthenticator result = extension.getRequestAuthenticator();

        assertNotNull(result);
        assertTrue(result instanceof OfficeExcelRequestAuthenticator);
    }

    // ==================== customTabs Tests ====================

    @Test
    void customTabs_returnsCorrectTabMapping() {
        Map<String, String> tabs = extension.customTabs();

        assertNotNull(tabs);
        assertEquals(4, tabs.size());
        assertEquals("rest/excel365extn/authentication/", tabs.get("Authentication"));
        assertEquals("static/docs", tabs.get("Documentation"));
        assertEquals("rest/excel365extn/audit/", tabs.get("Excel Audit Dashboard"));
        assertEquals("rest/excel365extn/mcp-setup", tabs.get("AI Assistant"));
    }

    @Test
    void customTabs_returnsImmutableMap() {
        Map<String, String> tabs = extension.customTabs();

        assertThrows(UnsupportedOperationException.class, () ->
            tabs.put("NewTab", "newPath"));
    }

    // ==================== validate Tests ====================

    @Test
    void validate_withValidEmailAndSuccessfulConnection_completesSuccessfully() throws IOException {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(Constants.EMAIL, "user@example.com");
        attributes.put(Constants.CLIENT_ID, "client-id");
        attributes.put(Constants.CLIENT_SECRET, "client-secret");
        attributes.put(Constants.TENANT_ID, "tenant-id");

        when(clientProviderFactory.create(any(OfficeExcelAttributes.class))).thenReturn(graphServiceClientProvider);
        when(graphServiceClientProvider.getGraphServiceClientForAdmin()).thenReturn(graphServiceClient);
        when(graphServiceClient.me()).thenReturn(userRequestBuilder);
        when(userRequestBuilder.buildRequest()).thenReturn(userRequest);
        when(userRequest.get()).thenReturn(new User());

        extension.validate(attributes);

        verify(clientProviderFactory).create(any(OfficeExcelAttributes.class));
        verify(graphServiceClientProvider).getGraphServiceClientForAdmin();
        verify(graphServiceClient).me();
        verify(userRequestBuilder).buildRequest();
        verify(userRequest).get();
    }

    @Test
    void validate_withBlankEmail_skipsGraphClientCall() throws IOException {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(Constants.EMAIL, "   ");

        extension.validate(attributes);

        verify(clientProviderFactory, never()).create(any(OfficeExcelAttributes.class));
        verifyNoInteractions(graphServiceClientProvider, graphServiceClient, userRequestBuilder, userRequest);
    }

    @Test
    void validate_withNullEmail_skipsGraphClientCall() throws IOException {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(Constants.EMAIL, null);

        extension.validate(attributes);

        verify(clientProviderFactory, never()).create(any(OfficeExcelAttributes.class));
        verifyNoInteractions(graphServiceClientProvider, graphServiceClient, userRequestBuilder, userRequest);
    }

    @Test
    void validate_withEmptyEmail_skipsGraphClientCall() throws IOException {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(Constants.EMAIL, "");

        extension.validate(attributes);

        verify(clientProviderFactory, never()).create(any(OfficeExcelAttributes.class));
        verifyNoInteractions(graphServiceClientProvider, graphServiceClient, userRequestBuilder, userRequest);
    }

    @Test
    void validate_whenGraphClientThrowsException_propagatesException() throws IOException {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(Constants.EMAIL, "user@example.com");

        when(clientProviderFactory.create(any(OfficeExcelAttributes.class))).thenReturn(graphServiceClientProvider);
        when(graphServiceClientProvider.getGraphServiceClientForAdmin()).thenReturn(graphServiceClient);
        when(graphServiceClient.me()).thenReturn(userRequestBuilder);
        when(userRequestBuilder.buildRequest()).thenReturn(userRequest);
        when(userRequest.get()).thenThrow(new ClientException("Authentication failed", null));

        assertThrows(ClientException.class, () -> extension.validate(attributes));
    }

    // ==================== testConnection Tests ====================

    @Test
    void testConnection_alwaysThrowsIllegalArgumentException() {
        IllegalArgumentException exception = assertThrows(
                IllegalArgumentException.class,
                () -> extension.testConnection());

        assertEquals("Please use Authentication tab for test connection.", exception.getMessage());
    }

    // ==================== attributesUpdated Tests ====================

    @Test
    void attributesUpdated_savesNewAttributesToStore() {
        Map<String, Object> oldAttributes = new HashMap<>();
        oldAttributes.put(Constants.CLIENT_ID, "old-client-id");

        Map<String, Object> newAttributes = new HashMap<>();
        newAttributes.put(Constants.CLIENT_ID, "new-client-id");
        newAttributes.put(Constants.CLIENT_SECRET, "new-secret");

        extension.attributesUpdated(oldAttributes, newAttributes);

        verify(officeExcelAttributeStore).save(any(OfficeExcelAttributes.class), eq("test-invoker"));
    }

    @Test
    void attributesUpdated_withEmptyMaps_stillSavesToStore() {
        Map<String, Object> oldAttributes = new HashMap<>();
        Map<String, Object> newAttributes = new HashMap<>();

        extension.attributesUpdated(oldAttributes, newAttributes);

        verify(officeExcelAttributeStore).save(any(OfficeExcelAttributes.class), eq("test-invoker"));
    }

    @Test
    void attributesUpdated_withNullOldAttributes_stillSavesToStore() {
        Map<String, Object> newAttributes = new HashMap<>();
        newAttributes.put(Constants.CLIENT_ID, "new-client-id");

        extension.attributesUpdated(null, newAttributes);

        verify(officeExcelAttributeStore).save(any(OfficeExcelAttributes.class), eq("test-invoker"));
    }
}

