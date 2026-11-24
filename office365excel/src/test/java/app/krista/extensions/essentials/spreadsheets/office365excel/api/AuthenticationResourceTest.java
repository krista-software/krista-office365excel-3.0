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

package app.krista.extensions.essentials.spreadsheets.office365excel.api;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.AuthenticationResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.authentication.AuthorizationListener;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.model.field.NamedValuedField;
import app.krista.sdk.msauth.MicrosoftAuthSdk;
import app.krista.sdk.msauth.extension.registry.ExtensionMetadata;
import app.krista.sdk.msauth.extension.registry.ExtensionRegistry;
import app.krista.sdk.msauth.token.ITokenStore;
import app.krista.sdk.msauth.token.TokenResult;
import com.google.gson.JsonObject;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.requests.GraphServiceClient;
import okhttp3.Request;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.MockedStatic;
import org.mockito.junit.jupiter.MockitoExtension;

import javax.ws.rs.BadRequestException;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Comprehensive test suite for AuthenticationResource.
 * Covers the Microsoft Auth SDK-based authentication flow (migrated from ScribeJava).
 */
@ExtendWith(MockitoExtension.class)
class AuthenticationResourceTest {

    @Mock
    private OfficeExcelAttributeStore officeExcelAttributeStore;
    @Mock
    private RefreshTokenStore refreshTokenStore;
    @Mock
    private GraphServiceClientProviderFactory providerFactory;
    @Mock
    private AuthorizationContext authorizationContext;
    @Mock
    private AuthorizationListener authorizationListener;
    @Mock
    private Invoker invoker;
    @Mock
    private RoutingInfo routingInfo;
    @Mock
    private Office365ExcelAuditService auditService;

    private AuthenticationResource resource;

    private static final String INVOKER_ID = "invoker-123";
    private static final String BASE_URL = "https://test.krista.app";

    @BeforeEach
    void setUp() {
        when(invoker.getInvokerId()).thenReturn(INVOKER_ID);
        when(invoker.getRoutingInfo()).thenReturn(routingInfo);
        when(routingInfo.getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE)).thenReturn(BASE_URL);
        // auditService is a lenient mock — no need to stub its methods
        lenient().doNothing().when(auditService).startRequest();
        lenient().doNothing().when(auditService).record(anyString(), anyString(), anyString(),
                anyString(), anyString(), any(), anyMap());
        resource = new AuthenticationResource(officeExcelAttributeStore, refreshTokenStore, providerFactory,
                authorizationContext, authorizationListener, invoker, auditService);
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_initializesAllFields() {
        assertNotNull(resource);
        verify(invoker).getInvokerId();
        verify(invoker).getRoutingInfo();
        verify(routingInfo).getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE);
    }

    // ==================== OAuth Callback Tests - NULL / Error Code Checks ====================

    @Test
    void getCallBackForV2Auth_nullCode_returnsFailureMessage() {
        String result = resource.getCallBackForV2Auth(null, "state", null);
        assertTrue(result.contains("Authentication Failed. Please re-authorize."));
    }

    @Test
    void getCallBackForV2Auth_nullState_throwsBadRequestException() {
        // null state → parts = [""], parts[0] is blank → BadRequestException
        BadRequestException ex = assertThrows(BadRequestException.class,
                () -> resource.getCallBackForV2Auth("code", null, null));
        assertTrue(ex.getMessage().contains("Invalid OAuth state"));
    }

    @Test
    void getCallbackForV3Auth_nullCode_returnsFailureMessage() {
        String result = resource.getCallbackForV3Auth(null, "state", null);
        assertTrue(result.contains("Authentication Failed. Please re-authorize."));
    }

    @Test
    void getCallbackForV3Auth_nullState_throwsBadRequestException() {
        BadRequestException ex = assertThrows(BadRequestException.class,
                () -> resource.getCallbackForV3Auth("code", null, null));
        assertTrue(ex.getMessage().contains("Invalid OAuth state"));
    }

    // ==================== OAuth Callback Tests - Invalid State ====================

    @Test
    void getCallBackForV2Auth_blankStateFirstPart_throwsBadRequestException() {
        // state = "#authCtx" → parts[0] is blank → BadRequestException
        BadRequestException ex = assertThrows(BadRequestException.class,
                () -> resource.getCallBackForV2Auth("code", "#authCtx", null));
        assertTrue(ex.getMessage().contains("Invalid OAuth state"));
    }

    // ==================== OAuth Callback Tests - PUBLIC Authentication ====================

    @Test
    void getCallBackForV2Auth_publicAuth_authenticatedContext_success() throws Exception {
        // state = "user@example.com#auth123#/forward/path"
        // authContextId = "auth123" (positional: parts[1] not starting with /)
        String state = "user@example.com#auth123#/forward/path";
        setupSuccessfulOAuthCallback(state, "auth123", true);

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(AuthenticationResource.USER_AUTHENTICATED_SUCCESSFULLY_PLEASE_PROCEED_WITH_REQUEST));
        verify(authorizationListener).authorized();
        verify(officeExcelAttributeStore).remove("auth123");
    }

    @Test
    void getCallBackForV2Auth_publicAuth_notAuthenticatedContext_success() throws Exception {
        // state = "user@example.com#auth123#/forward/path"
        String state = "user@example.com#auth123#/forward/path";
        setupSuccessfulOAuthCallback(state, "auth123", false);

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(Constants.USER_AUTHENTICATED_SUCCESSFULLY));
        verify(authorizationListener, never()).authorized();
        verify(officeExcelAttributeStore).remove("auth123");
    }

    // ==================== OAuth Callback Tests - PRIVATE Authentication ====================

    @Test
    void getCallBackForV2Auth_privateAuth_withAuthContextId_success() throws Exception {
        // state = "user@example.com#auth456"
        // authContextId = "auth456"
        String state = "user@example.com#auth456";
        setupSuccessfulOAuthCallback(state, "auth456", true);

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(AuthenticationResource.USER_AUTHENTICATED_SUCCESSFULLY_PLEASE_PROCEED_WITH_REQUEST));
        verify(authorizationListener).authorized();
        verify(officeExcelAttributeStore).remove("auth456");
    }

    @Test
    void getCallBackForV2Auth_privateAuth_noAuthContextId_success() throws Exception {
        // state = "user@example.com" → no authContextId → providerFactory.create() (no args)
        String state = "user@example.com";
        setupSuccessfulOAuthCallback(state, null, false);

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(Constants.USER_AUTHENTICATED_SUCCESSFULLY));
        verify(authorizationListener, never()).authorized();
        verify(officeExcelAttributeStore, never()).remove(anyString());
    }

    // ==================== OAuth Callback Tests - WS_CONTACT User ====================

    @Test
    void getCallBackForV2Auth_wsContactUser_skipsUserAccessCheck() throws Exception {
        // state = "wsContact123#auth789" (2 parts — no enriched-attrs save)
        // key = "wsContact123" (starts with WS_CONTACT) → authContextId = null
        // getCallbackProvider creates a fresh OfficeExcelAttributes (wsContact branch)
        String state = "wsContact123#auth789";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        TokenResult tokenResult = mock(TokenResult.class);

        // wsContact branch creates new OfficeExcelAttributes — match with any()
        when(providerFactory.create(any(OfficeExcelAttributes.class))).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any())).thenReturn(tokenResult);
        when(authorizationContext.isAuthenticated()).thenReturn(true);

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(AuthenticationResource.USER_AUTHENTICATED_SUCCESSFULLY_PLEASE_PROCEED_WITH_REQUEST));
        // wsContact key → user access check skipped entirely
        verify(clientProvider, never()).getAuthenticatedUserEmail();
        verify(authorizationListener).authorized();
        // authContextId = null for wsContact state → no store.remove call
        verify(officeExcelAttributeStore, never()).remove(anyString());
    }

    // ==================== wsContact MCP callback — enriched attrs saved ====================

    @Test
    void getCallbackForV3Auth_wsContactState_savesEnrichedAttrs() throws Exception {
        // state = "wsContact123#INV_abc#forwardPath"
        // key = "wsContact123" (starts with WS_CONTACT), clientKey = "INV_abc"
        String state = "wsContact123#INV_abc#/rest/excel365extn/v3/oauth/callback";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        TokenResult tokenResult = mock(TokenResult.class);

        // getCallbackProvider sees wsContact → creates attrs with null creds
        when(providerFactory.create(any(OfficeExcelAttributes.class))).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any())).thenReturn(tokenResult);
        // Token result returns a fake JWT — extractEmailFromAccessToken will fail gracefully
        when(tokenResult.getAccessToken()).thenReturn("header.e30K.sig"); // base64 of "{}"
        when(authorizationContext.isAuthenticated()).thenReturn(true);

        String result = resource.getCallbackForV3Auth("code", state, null);

        assertTrue(result.contains(AuthenticationResource.USER_AUTHENTICATED_SUCCESSFULLY_PLEASE_PROCEED_WITH_REQUEST));
        // Verify enriched attrs saved under wsContactId + suffix (not invokerId) — Rule 1 hermetic keyspace
        verify(officeExcelAttributeStore).save(any(OfficeExcelAttributes.class), eq("wsContact123_excel_attrs"));
    }

    // ==================== OAuth Callback Tests - User Access Denied ====================

    @Test
    void getCallBackForV2Auth_userAccessDenied_removesTokenAndReturnsError() throws Exception {
        // state = "user@example.com#auth123"
        // key = "user@example.com" (configuredEmail), authContextId = "auth123"
        // Access token contains a DIFFERENT email → mismatch → access denied
        String state = "user@example.com#auth123";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        TokenResult tokenResult = mock(TokenResult.class);

        when(officeExcelAttributeStore.load("auth123")).thenReturn(attrs);
        when(providerFactory.create(attrs)).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any())).thenReturn(tokenResult);
        // JWT payload decodes to {"preferred_username":"other@example.com"} — mismatches "user@example.com"
        when(tokenResult.getAccessToken()).thenReturn(
                "header.eyJwcmVmZXJyZWRfdXNlcm5hbWUiOiJvdGhlckBleGFtcGxlLmNvbSJ9.sig");

        String body = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(body.contains(Constants.UNAUTHORISED_USER));
        assertTrue(body.contains("user@example.com"));
        verify(refreshTokenStore).remove("user@example.com");
        verify(officeExcelAttributeStore).remove("auth123");
    }

    // ==================== OAuth Callback Tests - Exception Handling ====================

    @Test
    void getCallBackForV2Auth_runtimeException_withErrorMessage_returnsFormattedError() throws Exception {
        String state = "user@example.com#auth123";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);

        when(officeExcelAttributeStore.load("auth123")).thenReturn(attrs);
        when(providerFactory.create(attrs)).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any()))
                .thenThrow(new RuntimeException("Invalid authorization code"));

        String body = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(body.contains(Constants.ERROR_OCCURRED_DURING_AUTHORIZATION));
        assertTrue(body.contains("Invalid authorization code"));
        verify(officeExcelAttributeStore).remove("auth123");
    }

    @Test
    void getCallBackForV2Auth_runtimeException_withoutDetailedMessage_returnsGenericError() throws Exception {
        String state = "user@example.com#auth123";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);

        when(officeExcelAttributeStore.load("auth123")).thenReturn(attrs);
        when(providerFactory.create(attrs)).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any()))
                .thenThrow(new RuntimeException("Simple error"));

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(Constants.ERROR_OCCURRED_DURING_AUTHORIZATION));
        verify(officeExcelAttributeStore).remove("auth123");
    }

    @Test
    void getCallBackForV2Auth_interruptedException_setsInterruptFlagAndReturnsError() throws Exception {
        String state = "user@example.com#auth123";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);

        when(officeExcelAttributeStore.load("auth123")).thenReturn(attrs);
        when(providerFactory.create(attrs)).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        doAnswer(inv -> { throw new InterruptedException("Thread interrupted"); })
                .when(sdk).handleCallback(anyString(), anyString());

        String body = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(body.contains(Constants.ERROR_OCCURRED_DURING_AUTHORIZATION));
        assertTrue(body.contains("Thread interrupted"));
        assertTrue(Thread.currentThread().isInterrupted());
        Thread.interrupted(); // Clear interrupt flag so it does not affect subsequent tests
        verify(officeExcelAttributeStore).remove("auth123");
    }

    @Test
    void getCallBackForV2Auth_genericException_returnsError() throws Exception {
        String state = "user@example.com#auth123";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);

        when(officeExcelAttributeStore.load("auth123")).thenReturn(attrs);
        when(providerFactory.create(attrs)).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any()))
                .thenThrow(new RuntimeException("Unexpected error"));

        String body = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(body.contains(Constants.ERROR_OCCURRED_DURING_AUTHORIZATION));
        assertTrue(body.contains("Unexpected error"));
        verify(officeExcelAttributeStore).remove("auth123");
    }

    // ==================== Custom Tabs / Static Resources Tests ====================

    @Test
    void customTabs_emptyPath_returnsInputStreamOrNull() {
        // The resource may return null if the classpath resource is absent in unit test env.
        // Verify it does not throw.
        assertDoesNotThrow(() -> resource.customTabs(""));
    }

    @Test
    void customTabs_specificFile_doesNotThrow() {
        assertDoesNotThrow(() -> resource.customTabs("somefile.html"));
    }

    // ==================== Get Credentials Tests ====================

    @Test
    void getCredentials_publicAuth_ownMode_returnsCredentials() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getPublicAuthMode()).thenReturn("own");
        when(attrs.toMap()).thenReturn(Map.of("authType", Constants.PUBLIC, "email", "test@example.com"));
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        // publicAuthMode="own" (User Authentication) → credentials returned even with null authFlow
        String result = resource.getCredentials(Constants.PUBLIC, null);

        assertFalse(result.isEmpty());
        assertTrue(result.contains("authType"));
        assertTrue(result.contains(Constants.PUBLIC));
    }

    @Test
    void getCredentials_publicAuth_kristaMode_nullAuthFlow_returnsEmpty() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        // publicAuthMode=null defaults to "krista" (Quick Setup) — credentials must NOT leak
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, null);

        assertEquals("", result);
    }

    @Test
    void getCredentials_privateAuth_matchingAuthType_returnsCredentials() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PRIVATE);
        when(attrs.toMap()).thenReturn(Map.of("authType", Constants.PRIVATE, "clientId", "client123"));
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PRIVATE, null);

        assertFalse(result.isEmpty());
        assertTrue(result.contains("authType"));
        assertTrue(result.contains(Constants.PRIVATE));
    }

    @Test
    void getCredentials_mismatchedAuthType_returnsEmpty() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PRIVATE, null);

        assertEquals("", result);
    }

    @Test
    void getCredentials_noStoredAttributes_returnsEmpty() {
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(null);

        String result = resource.getCredentials(Constants.PUBLIC, null);

        assertEquals("", result);
    }

    @Test
    void getCredentials_quickSetupStored_userAuthFlowProbe_returnsEmpty() {
        // Regression test: Quick Setup (authFlow=AUTH_CODE_PUBLIC) must NOT be returned
        // when the SDK HTML probes for User Authentication (authFlow=AUTH_CODE_PRIVATE).
        // This was the bug that caused the auth tab to land on User Authentication
        // after Quick Setup was configured.
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, Constants.AUTH_FLOW_PRIVATE_USER);

        assertEquals("", result);
    }

    @Test
    void getCredentials_quickSetupStored_quickSetupFlowProbe_returnsCredentials() {
        // Quick Setup probe (authFlow=AUTH_CODE_PUBLIC) must match Quick Setup credentials.
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.toMap()).thenReturn(Map.of("authType", Constants.PUBLIC,
                "authFlow", Constants.AUTH_FLOW_PUBLIC, "email", "user@example.com"));
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, Constants.AUTH_FLOW_PUBLIC);

        assertFalse(result.isEmpty());
        assertTrue(result.contains(Constants.AUTH_FLOW_PUBLIC));
        // Quick Setup credentials must NOT expose clientId/clientSecret
        assertFalse(result.contains("clientId"));
        assertFalse(result.contains("clientSecret"));
    }

    @Test
    void getCredentials_userAuthStored_userAuthFlowProbe_returnsCredentials() {
        // User Authentication probe (authFlow=AUTH_CODE_PRIVATE) must match stored User Auth credentials.
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PRIVATE_USER);
        when(attrs.toMap()).thenReturn(Map.of("authType", Constants.PUBLIC,
                "authFlow", Constants.AUTH_FLOW_PRIVATE_USER,
                "email", "user@example.com", "clientId", "client123"));
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, Constants.AUTH_FLOW_PRIVATE_USER);

        assertFalse(result.isEmpty());
        assertTrue(result.contains(Constants.AUTH_FLOW_PRIVATE_USER));
        assertTrue(result.contains("clientId"));
    }

    // ==================== Backward Compat: publicAuthMode derivation from authFlow ====================

    @Test
    void getCredentials_publicAuth_nullPublicAuthMode_derivesFromAuthFlow_own() {
        // Backward compatibility: publicAuthMode is null, but authFlow=AUTH_CODE_PRIVATE -> derive "own"
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getPublicAuthMode()).thenReturn(null);
        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PRIVATE_USER);
        when(attrs.toMap()).thenReturn(Map.of("authType", Constants.PUBLIC,
                "authFlow", Constants.AUTH_FLOW_PRIVATE_USER,
                "email", "user@example.com", "clientId", "cid123"));
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, Constants.AUTH_FLOW_PRIVATE_USER);

        assertFalse(result.isEmpty());
        assertTrue(result.contains("clientId"));
    }

    @Test
    void getCredentials_publicAuth_nullPublicAuthMode_derivesFromAuthFlow_krista() {
        // Backward compatibility: publicAuthMode is null, authFlow=AUTH_CODE_PUBLIC -> derive "krista"
        // Quick Setup probe must return credentials (with credential masking)
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getPublicAuthMode()).thenReturn(null);
        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.toMap()).thenReturn(new HashMap<>(Map.of("authType", Constants.PUBLIC,
                "authFlow", Constants.AUTH_FLOW_PUBLIC, "email", "user@example.com",
                "clientId", "should-be-stripped", "clientSecret", "should-be-stripped")));
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, Constants.AUTH_FLOW_PUBLIC);

        assertFalse(result.isEmpty());
        // Belt-and-suspenders: even if toMap() included credentials, they're stripped
        assertFalse(result.contains("should-be-stripped"), "clientId must be stripped from Quick Setup response");
    }

    // ==================== Belt-and-suspenders credential stripping ====================

    @Test
    void getCredentials_publicAuth_kristaMode_beltsAndSuspenders_stripsCredentials() {
        // Explicit publicAuthMode="krista" with authFlow=AUTH_CODE_PUBLIC
        // Even if toMap() hypothetically includes credentials, getCredentials strips them
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getPublicAuthMode()).thenReturn("krista");
        when(attrs.toMap()).thenReturn(new HashMap<>(Map.of("authType", Constants.PUBLIC,
                "email", "user@example.com",
                "clientId", "leaked-cid", "clientSecret", "leaked-cs")));
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, Constants.AUTH_FLOW_PUBLIC);

        assertFalse(result.isEmpty());
        assertFalse(result.contains("leaked-cid"), "clientId must be stripped by belt-and-suspenders");
        assertFalse(result.contains("leaked-cs"), "clientSecret must be stripped by belt-and-suspenders");
    }

    // ==================== Admin Consent Callback ====================

    @Test
    void getCallBackForV2Auth_adminConsent_returnsSuccessMessage() {
        String body = resource.getCallBackForV2Auth(null, null, "True");
        assertTrue(body.contains("Admin consent granted successfully"));
    }

    @Test
    void getCallbackForV3Auth_adminConsent_returnsSuccessMessage() {
        String body = resource.getCallbackForV3Auth(null, null, "True");
        assertTrue(body.contains("Admin consent granted successfully"));
    }

    // ==================== Get Auth Key Tests ====================

    @Test
    void getAuthKey_publicAuthStored_returnsPublic() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getAuthKey();

        assertEquals("\"" + Constants.PUBLIC + "\"", result);
    }

    @Test
    void getAuthKey_privateAuthStored_returnsPrivate() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PRIVATE);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getAuthKey();

        assertEquals("\"" + Constants.PRIVATE + "\"", result);
    }

    @Test
    void getAuthKey_nullAuthType_returnsPublicAsDefault() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(null);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getAuthKey();

        assertEquals("\"" + Constants.PUBLIC + "\"", result);
    }

    @Test
    void getAuthKey_noStoredAttributes_returnsEmpty() {
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(null);

        String result = resource.getAuthKey();

        assertEquals("", result);
    }

    // ==================== Test Connection Tests ====================

    @Test
    void testConnection_publicAuth_success() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx1");
        when(providerFactory.create("ctx1")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);

        String result = resource.testConnection(payload);

        AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
        assertTrue(response.isSuccess());
        assertNull(response.getErrorMessage());
        assertNull(response.getUrl());
        verify(officeExcelAttributeStore).remove("ctx1");
    }

    @Test
    void testConnection_s2sClientCredentials_success_usesUsersDriveEndpoint() throws IOException {
        // PRIVATE authType → OfficeExcelAttributes.create sets authFlow=CLIENT_CREDENTIALS_SECRET
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PRIVATE);
        payload.addProperty(Constants.CLIENT_ID, "s2s-client");
        payload.addProperty(Constants.CLIENT_SECRET, "s2s-secret");
        payload.addProperty(Constants.TENANT_ID, "s2s-tenant");
        payload.addProperty(Constants.EMAIL, "svc@example.com");

        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        // RETURNS_DEEP_STUBS allows .users(email).drive().buildRequest().get() to succeed
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-s2s");
        when(providerFactory.create("ctx-s2s")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);

        String result = resource.testConnection(payload);

        AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
        assertTrue(response.isSuccess());
        assertNull(response.getErrorMessage());
        assertNull(response.getUrl());
        // S2S path uses .users().drive(), so .me() must never be called
        verify(graphClient, never()).me();
        verify(officeExcelAttributeStore).remove("ctx-s2s");
    }

    @Test
    void testConnection_s2sClientCredentials_mustAuthorizeException_returnsErrorDirectlyWithoutAuthUrl() throws Exception {
        // S2S has no OAuth flow — MustAuthorizeException must return a plain failure response
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PRIVATE);
        payload.addProperty(Constants.CLIENT_ID, "bad-client");
        payload.addProperty(Constants.CLIENT_SECRET, "bad-secret");
        payload.addProperty(Constants.TENANT_ID, "bad-tenant");
        payload.addProperty(Constants.EMAIL, "svc@example.com");

        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);

        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-s2s-err");
        when(providerFactory.create("ctx-s2s-err")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException(
                        "Invalid credentials. Please verify your Client ID, Client Secret, and Tenant ID.",
                        Collections.emptyList()));

        String result = resource.testConnection(payload);

        AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
        assertFalse(response.isSuccess());
        assertTrue(response.getErrorMessage().contains("Invalid credentials"));
        // Must NOT generate an OAuth redirect URL for S2S
        assertNull(response.getUrl());
        verify(officeExcelAttributeStore).remove("ctx-s2s-err");
    }

    @Test
    void testConnection_nullEmail_successWithNoTriggeredByInMeta() throws IOException {
        // email is null → meta.put("triggeredBy", ...) branch is skipped
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        // no EMAIL property → getEmail() returns null

        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-null-email");
        when(providerFactory.create("ctx-null-email")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);

        String result = resource.testConnection(payload);

        AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
        assertTrue(response.isSuccess());
        verify(officeExcelAttributeStore).remove("ctx-null-email");
    }

    @Test
    void testConnection_graphServiceException_returnsFailureWithGenericMessage() throws Exception {
        // GraphServiceException → caught, returns failure with generic message, authUrl stays null → finally removes ctx
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceException graphEx = mock(GraphServiceException.class);
        when(graphEx.getMessage()).thenReturn("404 Not Found");

        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-graph-ex");
        when(providerFactory.create("ctx-graph-ex")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenThrow(graphEx);

        String result = resource.testConnection(payload);

        AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
        assertFalse(response.isSuccess());
        assertEquals("An error occurred during test connection.", response.getErrorMessage());
        assertNull(response.getUrl());
        // authUrl is null → finally block must remove the auth context
        verify(officeExcelAttributeStore).remove("ctx-graph-ex");
    }

    @Test
    void testConnection_ioException_throwsRuntimeException() throws Exception {
        // IOException → wrapped in RuntimeException and re-thrown, finally still removes ctx
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);

        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-io");
        when(providerFactory.create("ctx-io")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenAnswer(inv -> { throw new IOException("IO error"); });

        RuntimeException ex = assertThrows(RuntimeException.class, () -> resource.testConnection(payload));
        assertInstanceOf(IOException.class, ex.getCause());
        verify(officeExcelAttributeStore).remove("ctx-io");
    }

    @Test
    void testConnection_mustAuthorizeException_delegated_urlHasFilesReadWrite_urlHasPrompt_replacesPromptWithConsent() throws Exception {
        // SDK URL already has Files.ReadWrite → no scope override
        // SDK URL already has prompt= → replaced with prompt=consent
        // authUrl is set → finally block must NOT remove auth context
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        NamedValuedField userIdField = mock(NamedValuedField.class);
        when(userIdField.getValue()).thenReturn("user@example.com");
        when(userIdField.getName()).thenReturn("userId");

        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.getForwardPath()).thenReturn("");
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.buildAuthorizationUrl(anyString()))
                .thenReturn("https://login.microsoftonline.com/auth?scope=openid%20Files.ReadWrite&prompt=select_account");
        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-mae-1");
        when(providerFactory.create("ctx-mae-1")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Auth required", List.of(userIdField)));

        try (MockedStatic<OfficeExcelAttributes> mockedStatic = mockStatic(OfficeExcelAttributes.class)) {
            mockedStatic.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attrs);

            String result = resource.testConnection(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertNotNull(response.getUrl());
            assertTrue(response.getUrl().contains("Files.ReadWrite"));
            assertTrue(response.getUrl().contains("prompt=consent"));
            assertFalse(response.getUrl().contains("prompt=select_account"));
            // authUrl != null → finally must NOT remove the auth context
            verify(officeExcelAttributeStore, never()).remove("ctx-mae-1");
        }
    }

    @Test
    void testConnection_mustAuthorizeException_delegated_urlHasFilesReadWrite_noPromptInUrl_appendsPromptConsent() throws Exception {
        // SDK URL has Files.ReadWrite → no scope override
        // SDK URL has no prompt= → &prompt=consent appended
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        NamedValuedField userIdField = mock(NamedValuedField.class);
        when(userIdField.getValue()).thenReturn("user@example.com");
        when(userIdField.getName()).thenReturn("userId");

        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.getForwardPath()).thenReturn("");
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.buildAuthorizationUrl(anyString()))
                .thenReturn("https://login.microsoftonline.com/auth?scope=openid%20Files.ReadWrite");
        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-mae-2");
        when(providerFactory.create("ctx-mae-2")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Auth required", List.of(userIdField)));

        try (MockedStatic<OfficeExcelAttributes> mockedStatic = mockStatic(OfficeExcelAttributes.class)) {
            mockedStatic.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attrs);

            String result = resource.testConnection(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertNotNull(response.getUrl());
            assertTrue(response.getUrl().contains("Files.ReadWrite"));
            assertTrue(response.getUrl().contains("prompt=consent"));
            verify(officeExcelAttributeStore, never()).remove("ctx-mae-2");
        }
    }

    @Test
    void testConnection_mustAuthorizeException_delegated_scopeMissingFilesReadWrite_urlHasPrompt_overridesScopeAndReplacesPrompt() throws Exception {
        // SDK URL has wrong scopes (no Files.ReadWrite) → scope is overridden with correct Excel scopes
        // SDK URL has prompt= → replaced with prompt=consent
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        NamedValuedField userIdField = mock(NamedValuedField.class);
        when(userIdField.getValue()).thenReturn("user@example.com");
        when(userIdField.getName()).thenReturn("userId");

        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.getForwardPath()).thenReturn("");
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        // SDK generated URL with wrong mail scopes and prompt=select_account
        when(sdk.buildAuthorizationUrl(anyString()))
                .thenReturn("https://login.microsoftonline.com/auth?scope=openid%20Mail.Send&prompt=select_account");
        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-mae-3");
        when(providerFactory.create("ctx-mae-3")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Auth required", List.of(userIdField)));

        try (MockedStatic<OfficeExcelAttributes> mockedStatic = mockStatic(OfficeExcelAttributes.class)) {
            mockedStatic.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attrs);

            String result = resource.testConnection(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertNotNull(response.getUrl());
            // Scope was overridden to include Files.ReadWrite
            assertTrue(response.getUrl().contains("Files.ReadWrite"));
            // prompt=select_account was replaced with prompt=consent
            assertTrue(response.getUrl().contains("prompt=consent"));
            assertFalse(response.getUrl().contains("prompt=select_account"));
            verify(officeExcelAttributeStore, never()).remove("ctx-mae-3");
        }
    }

    @Test
    void testConnection_mustAuthorizeException_delegated_scopeMissingFilesReadWrite_noPromptInUrl_overridesScopeAndAppendsPrompt() throws Exception {
        // SDK URL has wrong scopes (no Files.ReadWrite) → scope is overridden
        // SDK URL has no prompt= → &prompt=consent appended
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        NamedValuedField userIdField = mock(NamedValuedField.class);
        when(userIdField.getValue()).thenReturn("user@example.com");
        when(userIdField.getName()).thenReturn("userId");

        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.getForwardPath()).thenReturn("");
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        // SDK generated URL with wrong mail scopes, no prompt parameter
        when(sdk.buildAuthorizationUrl(anyString()))
                .thenReturn("https://login.microsoftonline.com/auth?scope=openid%20Mail.Send");
        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-mae-4");
        when(providerFactory.create("ctx-mae-4")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Auth required", List.of(userIdField)));

        try (MockedStatic<OfficeExcelAttributes> mockedStatic = mockStatic(OfficeExcelAttributes.class)) {
            mockedStatic.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attrs);

            String result = resource.testConnection(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertNotNull(response.getUrl());
            assertTrue(response.getUrl().contains("Files.ReadWrite"));
            assertTrue(response.getUrl().contains("prompt=consent"));
            verify(officeExcelAttributeStore, never()).remove("ctx-mae-4");
        }
    }

    // ==================== Save Credentials Tests ====================

    @Test
    void saveCredentials_publicAuth_success() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attributes = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(providerFactory.create(attributes)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);
        when(officeExcelAttributeStore.save(attributes)).thenReturn("ctx1");
        when(officeExcelAttributeStore.save(attributes, INVOKER_ID)).thenReturn(true);

        try (MockedStatic<OfficeExcelAttributes> mockedAttributes = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttributes.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attributes);

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertTrue(response.isSuccess());
            assertNull(response.getErrorMessage());
            verify(officeExcelAttributeStore).remove("ctx1");
        }
    }

    @Test
    void saveCredentials_privateAuth_success() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PRIVATE);
        payload.addProperty(Constants.CLIENT_ID, "client123");
        payload.addProperty(Constants.CLIENT_SECRET, "secret456");

        OfficeExcelAttributes attributes = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(providerFactory.create(attributes)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);
        when(officeExcelAttributeStore.save(attributes)).thenReturn("ctx2");
        when(officeExcelAttributeStore.save(attributes, INVOKER_ID)).thenReturn(true);

        try (MockedStatic<OfficeExcelAttributes> mockedAttributes = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttributes.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attributes);

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertTrue(response.isSuccess());
            verify(officeExcelAttributeStore).remove("ctx2");
        }
    }

    @Test
    void saveCredentials_saveFails_returnsError() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attributes = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(providerFactory.create(attributes)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);
        when(officeExcelAttributeStore.save(attributes)).thenReturn("ctx3");
        when(officeExcelAttributeStore.save(attributes, INVOKER_ID)).thenReturn(false);

        try (MockedStatic<OfficeExcelAttributes> mockedAttributes = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttributes.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attributes);

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertEquals(Constants.FAILED_TO_SAVE_ATTRIBUTES, response.getErrorMessage());
            verify(officeExcelAttributeStore).remove("ctx3");
        }
    }

    @Test
    void saveCredentials_s2sClientCredentials_success_usesUsersDriveEndpoint() throws IOException {
        // S2S (PRIVATE) → authFlow=CLIENT_CREDENTIALS_SECRET → uses .users(email).drive()
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PRIVATE);
        payload.addProperty(Constants.CLIENT_ID, "s2s-client");
        payload.addProperty(Constants.CLIENT_SECRET, "s2s-secret");
        payload.addProperty(Constants.TENANT_ID, "s2s-tenant");
        payload.addProperty(Constants.EMAIL, "svc@example.com");

        OfficeExcelAttributes attributes = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(attributes.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_CLIENT_CREDS);
        when(attributes.getEmail()).thenReturn("svc@example.com");
        when(providerFactory.create(attributes)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);
        when(officeExcelAttributeStore.save(attributes)).thenReturn("ctx-s2s-save");
        when(officeExcelAttributeStore.save(attributes, INVOKER_ID)).thenReturn(true);

        try (MockedStatic<OfficeExcelAttributes> mockedAttributes = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttributes.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attributes);

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertTrue(response.isSuccess());
            assertNull(response.getErrorMessage());
            // S2S path uses .users().drive(), so .me() must never be called
            verify(graphClient, never()).me();
            verify(officeExcelAttributeStore).remove("ctx-s2s-save");
        }
    }

    @Test
    void saveCredentials_s2sClientCredentials_mustAuthorizeException_returnsFailureResponseNotRethrow() throws IOException {
        // S2S MustAuthorizeException must return AuthenticationResponse(false, msg, null)
        // and must NOT re-throw (which would create an inappropriate OAuth redirect for a non-OAuth flow)
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PRIVATE);
        payload.addProperty(Constants.CLIENT_ID, "bad-client");
        payload.addProperty(Constants.CLIENT_SECRET, "bad-secret");
        payload.addProperty(Constants.TENANT_ID, "bad-tenant");
        payload.addProperty(Constants.EMAIL, "svc@example.com");

        OfficeExcelAttributes attributes = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);

        when(attributes.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_CLIENT_CREDS);
        when(attributes.getEmail()).thenReturn("svc@example.com");
        when(providerFactory.create(attributes)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("S2S credential validation failed.", Collections.emptyList()));
        when(officeExcelAttributeStore.save(attributes)).thenReturn("ctx-s2s-mae");

        try (MockedStatic<OfficeExcelAttributes> mockedAttributes = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttributes.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attributes);

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertEquals("S2S credential validation failed.", response.getErrorMessage());
            // Must not include an auth URL — S2S has no OAuth redirect
            assertNull(response.getUrl());
            verify(officeExcelAttributeStore).remove("ctx-s2s-mae");
        }
    }

    @Test
    void saveCredentials_exceptionDuringValidation_returnsError() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attributes = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);

        when(providerFactory.create(attributes)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new RuntimeException("Validation failed"));
        when(officeExcelAttributeStore.save(attributes)).thenReturn("ctx4");

        try (MockedStatic<OfficeExcelAttributes> mockedAttributes = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttributes.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attributes);

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertEquals(Constants.FAILED_TO_SAVE_ATTRIBUTES, response.getErrorMessage());
            verify(officeExcelAttributeStore).remove("ctx4");
        }
    }

    // ==================== Private Method Tests - getAuthContextId ====================

    @Test
    void getAuthContextId_nullState_returnsNull() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        assertNull(method.invoke(resource, (String) null));
    }

    @Test
    void getAuthContextId_emptyState_returnsNull() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        assertNull(method.invoke(resource, ""));
    }

    @Test
    void getAuthContextId_singlePart_returnsNull() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        assertNull(method.invoke(resource, "user@example.com"));
    }

    @Test
    void getAuthContextId_twoPartsWithAuthContextId_returnsAuthContextId() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        assertEquals("ctx123", method.invoke(resource, "user@example.com#ctx123"));
    }

    @Test
    void getAuthContextId_twoPartsWithForwardPath_returnsNull() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        assertNull(method.invoke(resource, "user@example.com#/forward/path"));
    }

    @Test
    void getAuthContextId_threePartsPublicAuth_returnsAuthContextId() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        assertEquals("ctx456", method.invoke(resource, "user@example.com#ctx456#/forward/path"));
    }

    @Test
    void getAuthContextId_blankAuthContextId_returnsNull() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        assertNull(method.invoke(resource, "user@example.com# "));
    }

    // ==================== OAuth callback — blank authenticated email ====================

    @Test
    void getCallBackForV2Auth_blankAuthenticatedEmail_skipsEmailCheckAndSucceeds() throws Exception {
        // Covers: authenticatedEmail != null && authenticatedEmail.isBlank() = true
        // (the blank-email branch of the compound condition on line 171)
        // JWT has preferred_username = "" (empty string) → isBlank() = true → inner check skipped
        String state = "user@example.com#auth123";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        TokenResult tokenResult = mock(TokenResult.class);

        when(officeExcelAttributeStore.load("auth123")).thenReturn(attrs);
        when(providerFactory.create(attrs)).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any())).thenReturn(tokenResult);

        // JWT with blank preferred_username — authenticatedEmail = "" → isBlank() = true → skip rejection
        String jwtPayload = "{\"preferred_username\":\"\"}";
        String encodedPayload = java.util.Base64.getUrlEncoder().withoutPadding()
                .encodeToString(jwtPayload.getBytes());
        when(tokenResult.getAccessToken()).thenReturn("header." + encodedPayload + ".sig");
        when(authorizationContext.isAuthenticated()).thenReturn(false);

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(Constants.USER_AUTHENTICATED_SUCCESSFULLY));
        verify(officeExcelAttributeStore).remove("auth123");
    }

    // ==================== getCallbackProvider — null key branch via reflection ====================

    @Test
    void getCallbackProvider_nullKey_usesAuthContextIdBranch() throws Exception {
        // Covers: key == null → key != null is false → short-circuit → check authContextId
        // This exercises the unreachable-in-production branch where key is null.
        Method method = AuthenticationResource.class.getDeclaredMethod(
                "getCallbackProvider", String.class, String.class);
        method.setAccessible(true);

        GraphServiceClientProvider expectedProvider = mock(GraphServiceClientProvider.class);
        when(providerFactory.create()).thenReturn(expectedProvider);

        // key = null, authContextId = null → providerFactory.create() (no-args) path
        GraphServiceClientProvider result =
                (GraphServiceClientProvider) method.invoke(resource, (String) null, (String) null);

        assertSame(expectedProvider, result);
        verify(providerFactory).create();
    }

    // ==================== getCredentials — missing branches ====================

    @Test
    void getCredentials_nullAuthTypeParam_returnsEmpty() {
        // loadedAttributes is non-null but authType param is null → short-circuit returns ""
        // getAuthType() is never reached due to || short-circuit, so no stub needed
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(null, null);

        assertEquals("", result);
    }

    @Test
    void getCredentials_ownModeStoredProbedWithQuickSetupFlow_returnsEmpty() {
        // Regression: User Authentication stored ("own" mode) must NOT be returned
        // when the SDK probes with AUTH_CODE_PUBLIC (Quick Setup) authFlow.
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(attrs.getPublicAuthMode()).thenReturn("own");
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);

        String result = resource.getCredentials(Constants.PUBLIC, Constants.AUTH_FLOW_PUBLIC);

        assertEquals("", result);
    }

    // ==================== extractEmailFromAccessToken — private method via reflection ====================

    @Test
    void extractEmailFromAccessToken_tokenWithNoDots_returnsNull() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("extractEmailFromAccessToken", String.class);
        method.setAccessible(true);

        // Token with no dots → jwtParts.length == 1 (< 2) → returns null without entering catch
        assertNull(method.invoke(resource, "nodots"));
    }

    @Test
    void extractEmailFromAccessToken_tokenWithUpnClaim_andPaddingNeeded_returnsUpn() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("extractEmailFromAccessToken", String.class);
        method.setAccessible(true);

        // {"upn":"u@e.com"} = 17 bytes → base64url without padding = 23 chars → 23 % 4 = 3 → mod > 0
        // Covers: mod > 0 (padding branch) AND upn claim (no preferred_username)
        String payload = "{\"upn\":\"u@e.com\"}";
        String encodedPayload = java.util.Base64.getUrlEncoder().withoutPadding()
                .encodeToString(payload.getBytes());
        String token = "header." + encodedPayload + ".sig";

        assertEquals("u@e.com", method.invoke(resource, token));
    }

    @Test
    void extractEmailFromAccessToken_tokenWithEmailClaimOnly_returnsEmail() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("extractEmailFromAccessToken", String.class);
        method.setAccessible(true);

        // No preferred_username, no upn — only "email" claim
        String payload = "{\"email\":\"em@test.com\"}";
        String encodedPayload = java.util.Base64.getUrlEncoder().withoutPadding()
                .encodeToString(payload.getBytes());
        String token = "header." + encodedPayload + ".sig";

        assertEquals("em@test.com", method.invoke(resource, token));
    }

    // ==================== createTestConnectionResponse — private method via reflection ====================

    @Test
    void createTestConnectionResponse_nullMessage_nullUrl_returnsUnknownError() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod(
                "createTestConnectionResponse", boolean.class, String.class, String.class);
        method.setAccessible(true);

        // isSuccess=false, url=null, errorMessage=null → Objects.requireNonNullElse → "Unknown Error"
        String result = (String) method.invoke(resource, false, null, null);

        AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
        assertFalse(response.isSuccess());
        assertEquals("Unknown Error", response.getErrorMessage());
        assertNull(response.getUrl());
    }

    // ==================== createStateParameter — additional branches via testConnection ====================

    @Test
    void testConnection_mustAuthorizeException_stateIncludesAuthContextId_fromCauseDetails() throws Exception {
        // Covers: authContextIdField.isPresent() == true branch in createStateParameter
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);

        NamedValuedField userIdField = mock(NamedValuedField.class);
        when(userIdField.getName()).thenReturn("userId");
        when(userIdField.getValue()).thenReturn("user@example.com");

        NamedValuedField authCtxField = mock(NamedValuedField.class);
        when(authCtxField.getName()).thenReturn(Constants.AUTH_CONTEXT_ID);
        when(authCtxField.getValue()).thenReturn("ctx-from-cause");

        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.getForwardPath()).thenReturn("");
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.buildAuthorizationUrl(anyString()))
                .thenReturn("https://login.microsoftonline.com/auth?scope=openid%20Files.ReadWrite&prompt=select_account");
        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-with-authctx");
        when(providerFactory.create("ctx-with-authctx")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Auth required",
                        List.of(userIdField, authCtxField)));

        try (MockedStatic<OfficeExcelAttributes> mockedStatic = mockStatic(OfficeExcelAttributes.class)) {
            mockedStatic.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attrs);

            String result = resource.testConnection(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertNotNull(response.getUrl());
            // authUrl was set → finally must NOT remove context
            verify(officeExcelAttributeStore, never()).remove("ctx-with-authctx");
        }
    }

    @Test
    void testConnection_mustAuthorizeException_withNonEmptyForwardPath_stateIncludesPath() throws Exception {
        // Covers: !forwardPath.isEmpty() == true branch in createStateParameter
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        NamedValuedField userIdField = mock(NamedValuedField.class);
        when(userIdField.getName()).thenReturn("userId");
        when(userIdField.getValue()).thenReturn("user@example.com");

        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(attrs.getForwardPath()).thenReturn("/rest/excel365extn/v3/oauth/callback"); // non-empty
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.buildAuthorizationUrl(anyString()))
                .thenReturn("https://login.microsoftonline.com/auth?scope=openid%20Files.ReadWrite&prompt=select_account");
        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-fwd");
        when(providerFactory.create("ctx-fwd")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Auth required", List.of(userIdField)));

        try (MockedStatic<OfficeExcelAttributes> mockedStatic = mockStatic(OfficeExcelAttributes.class)) {
            mockedStatic.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attrs);

            String result = resource.testConnection(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertNotNull(response.getUrl());
            verify(officeExcelAttributeStore, never()).remove("ctx-fwd");
        }
    }

    @Test
    void testConnection_mustAuthorizeException_privateUserAuthFlow_stateIncludesCallbackEndpoint() throws Exception {
        // Covers: AUTH_FLOW_PRIVATE_USER path in createStateParameter (empty forwardPath)
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        NamedValuedField userIdField = mock(NamedValuedField.class);
        when(userIdField.getName()).thenReturn("userId");
        when(userIdField.getValue()).thenReturn("user@example.com");

        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PRIVATE_USER); // own Azure AD app
        when(attrs.getForwardPath()).thenReturn(""); // empty → falls through to callbackEndPoint
        when(attrs.getCallbackEndPoint()).thenReturn("/rest/excel365extn/callback");
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.buildAuthorizationUrl(anyString()))
                .thenReturn("https://login.microsoftonline.com/auth?scope=openid%20Files.ReadWrite&prompt=select_account");
        when(providerFactory.createAttributes(any(OfficeExcelAttributes.class))).thenReturn("ctx-pvt");
        when(providerFactory.create("ctx-pvt")).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Auth required", List.of(userIdField)));

        try (MockedStatic<OfficeExcelAttributes> mockedStatic = mockStatic(OfficeExcelAttributes.class)) {
            mockedStatic.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attrs);

            String result = resource.testConnection(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertFalse(response.isSuccess());
            assertNotNull(response.getUrl());
            verify(officeExcelAttributeStore, never()).remove("ctx-pvt");
        }
    }

    // ==================== createStateParameter — null cause branch via reflection ====================

    @Test
    void createStateParameter_nullCause_withEmailId_returnsEmailId() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod(
                "createStateParameter", MustAuthorizeException.class, OfficeExcelAttributes.class);
        method.setAccessible(true);

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getEmailId()).thenReturn("user@example.com");
        when(attrs.getForwardPath()).thenReturn("");
        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);

        String state = (String) method.invoke(resource, (MustAuthorizeException) null, attrs);

        assertEquals("user@example.com", state);
    }

    @Test
    void createStateParameter_nullCause_withNullEmailId_usesTestUserFallback() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod(
                "createStateParameter", MustAuthorizeException.class, OfficeExcelAttributes.class);
        method.setAccessible(true);

        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(attrs.getEmailId()).thenReturn(null); // null → "test-user" fallback
        when(attrs.getForwardPath()).thenReturn("");
        when(attrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);

        String state = (String) method.invoke(resource, (MustAuthorizeException) null, attrs);

        assertEquals("test-user", state);
    }

    // ==================== getAuthContextId — AUTH_CONTEXT_ID prefix path via reflection ====================

    @Test
    void getAuthContextId_stateContainsAuthContextIdPrefix_withHashAfter_returnsId() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        // AUTH_CONTEXT_ID = "authContextId". state contains it → use prefix extraction
        // "user@example.com#authContextIdXYZ#/fwd" → authContextIdToEnd = "authContextIdXYZ#/fwd"
        // hashIndex != -1 → return "authContextIdXYZ"
        String state = "user@example.com#authContextIdXYZ#/fwd";
        assertEquals("authContextIdXYZ", method.invoke(resource, state));
    }

    @Test
    void getAuthContextId_stateContainsAuthContextIdPrefix_withoutHash_returnsFullId() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("getAuthContextId", String.class);
        method.setAccessible(true);

        // "user@example.com#authContextIdXYZ" → authContextIdToEnd = "authContextIdXYZ"
        // hashIndex == -1 → return full "authContextIdXYZ"
        String state = "user@example.com#authContextIdXYZ";
        assertEquals("authContextIdXYZ", method.invoke(resource, state));
    }

    // ==================== createTokenStore inner class via reflection ====================

    @Test
    void createTokenStore_saveSucceeds_delegatesToRefreshTokenStore() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("createTokenStore");
        method.setAccessible(true);
        ITokenStore tokenStore = (ITokenStore) method.invoke(resource);

        tokenStore.save("key1", "val1");

        verify(refreshTokenStore).put("key1", "val1");
    }

    @Test
    void createTokenStore_saveIOException_wrapsAsRuntimeException() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("createTokenStore");
        method.setAccessible(true);
        ITokenStore tokenStore = (ITokenStore) method.invoke(resource);

        doThrow(new IOException("put failed")).when(refreshTokenStore).put(anyString(), anyString());

        RuntimeException ex = assertThrows(RuntimeException.class, () -> tokenStore.save("k", "v"));
        assertInstanceOf(IOException.class, ex.getCause());
    }

    @Test
    void createTokenStore_getSucceeds_delegatesToRefreshTokenStore() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("createTokenStore");
        method.setAccessible(true);
        ITokenStore tokenStore = (ITokenStore) method.invoke(resource);

        when(refreshTokenStore.get("key1")).thenReturn("val1");

        assertEquals("val1", tokenStore.get("key1"));
    }

    @Test
    void createTokenStore_getIOException_wrapsAsRuntimeException() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("createTokenStore");
        method.setAccessible(true);
        ITokenStore tokenStore = (ITokenStore) method.invoke(resource);

        doThrow(new IOException("get failed")).when(refreshTokenStore).get(anyString());

        RuntimeException ex = assertThrows(RuntimeException.class, () -> tokenStore.get("k"));
        assertInstanceOf(IOException.class, ex.getCause());
    }

    @Test
    void createTokenStore_deleteSucceeds_delegatesToRefreshTokenStore() throws Exception {
        Method method = AuthenticationResource.class.getDeclaredMethod("createTokenStore");
        method.setAccessible(true);
        ITokenStore tokenStore = (ITokenStore) method.invoke(resource);

        tokenStore.delete("key1");

        verify(refreshTokenStore).remove("key1");
    }

    // ==================== getExtensionMetadata Tests ====================

    @Test
    void getExtensionMetadata_extensionFound_returnsUiJson() {
        ExtensionMetadata mockMetadata = mock(ExtensionMetadata.class);
        when(mockMetadata.toUiJson()).thenReturn("{\"id\":\"test-ext\"}");

        try (MockedStatic<ExtensionRegistry> registry = mockStatic(ExtensionRegistry.class)) {
            registry.when(ExtensionRegistry::initialize).thenAnswer(inv -> null);
            registry.when(() -> ExtensionRegistry.get("test-ext")).thenReturn(Optional.of(mockMetadata));

            String result = resource.getExtensionMetadata("test-ext");

            assertEquals("{\"id\":\"test-ext\"}", result);
        }
    }

    @Test
    void getExtensionMetadata_extensionNotFound_returnsErrorJson() {
        try (MockedStatic<ExtensionRegistry> registry = mockStatic(ExtensionRegistry.class)) {
            registry.when(ExtensionRegistry::initialize).thenAnswer(inv -> null);
            registry.when(() -> ExtensionRegistry.get("unknown-ext")).thenReturn(Optional.empty());

            String result = resource.getExtensionMetadata("unknown-ext");

            assertTrue(result.contains("Extension not found: unknown-ext"));
        }
    }

    // ==================== saveCredentials — old token cleanup branches ====================

    @Test
    void saveCredentials_oldCredentialsWithDifferentTokenKey_removesOldToken() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "new@example.com");

        OfficeExcelAttributes newAttrs = mock(OfficeExcelAttributes.class);
        OfficeExcelAttributes oldAttrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(newAttrs.getEmail()).thenReturn("new@example.com");
        when(newAttrs.getClientId()).thenReturn("new-cid");
        when(newAttrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(newAttrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(oldAttrs.getEmail()).thenReturn("old@example.com");
        when(oldAttrs.getClientId()).thenReturn("old-cid");
        when(oldAttrs.getAuthType()).thenReturn(Constants.PUBLIC);

        when(providerFactory.create(newAttrs)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);
        when(officeExcelAttributeStore.save(newAttrs)).thenReturn("ctx-old-diff");
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(oldAttrs);
        when(officeExcelAttributeStore.save(newAttrs, INVOKER_ID)).thenReturn(true);

        try (MockedStatic<OfficeExcelAttributes> mockedAttrs = mockStatic(OfficeExcelAttributes.class);
             MockedStatic<MicrosoftAuthSdk> mockedSdk = mockStatic(MicrosoftAuthSdk.class)) {
            mockedAttrs.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(newAttrs);
            mockedSdk.when(() -> MicrosoftAuthSdk.buildTokenKey("old@example.com", "old-cid", Constants.PUBLIC))
                    .thenReturn("old-token-key");
            mockedSdk.when(() -> MicrosoftAuthSdk.buildTokenKey("new@example.com", "new-cid", Constants.PUBLIC))
                    .thenReturn("new-token-key");

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertTrue(response.isSuccess());
            verify(refreshTokenStore).remove("old-token-key");
        }
    }

    @Test
    void saveCredentials_oldCredentialsWithSameTokenKey_doesNotRemoveToken() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes newAttrs = mock(OfficeExcelAttributes.class);
        OfficeExcelAttributes oldAttrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(newAttrs.getEmail()).thenReturn("user@example.com");
        when(newAttrs.getClientId()).thenReturn("cid");
        when(newAttrs.getAuthType()).thenReturn(Constants.PUBLIC);
        when(newAttrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(oldAttrs.getEmail()).thenReturn("user@example.com");
        when(oldAttrs.getClientId()).thenReturn("cid");
        when(oldAttrs.getAuthType()).thenReturn(Constants.PUBLIC);

        when(providerFactory.create(newAttrs)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);
        when(officeExcelAttributeStore.save(newAttrs)).thenReturn("ctx-same-key");
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(oldAttrs);
        when(officeExcelAttributeStore.save(newAttrs, INVOKER_ID)).thenReturn(true);

        try (MockedStatic<OfficeExcelAttributes> mockedAttrs = mockStatic(OfficeExcelAttributes.class);
             MockedStatic<MicrosoftAuthSdk> mockedSdk = mockStatic(MicrosoftAuthSdk.class)) {
            mockedAttrs.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(newAttrs);
            // Both keys are identical → no removal
            mockedSdk.when(() -> MicrosoftAuthSdk.buildTokenKey("user@example.com", "cid", Constants.PUBLIC))
                    .thenReturn("same-token-key");

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertTrue(response.isSuccess());
            verify(refreshTokenStore, never()).remove("same-token-key");
        }
    }

    @Test
    void saveCredentials_oldCredentialsWithNullEmail_skipsTokenCleanup() throws IOException {
        // oldAttributes.getEmail() == null → condition is false → skip cleanup block entirely
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "new@example.com");

        OfficeExcelAttributes newAttrs = mock(OfficeExcelAttributes.class);
        OfficeExcelAttributes oldAttrs = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        GraphServiceClient<?> graphClient = mock(GraphServiceClient.class, RETURNS_DEEP_STUBS);

        when(newAttrs.getEmail()).thenReturn("new@example.com");
        when(newAttrs.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC);
        when(oldAttrs.getEmail()).thenReturn(null); // null email → skip token cleanup

        when(providerFactory.create(newAttrs)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin()).thenReturn((GraphServiceClient<Request>) graphClient);
        when(officeExcelAttributeStore.save(newAttrs)).thenReturn("ctx-null-old-email");
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(oldAttrs);
        when(officeExcelAttributeStore.save(newAttrs, INVOKER_ID)).thenReturn(true);

        try (MockedStatic<OfficeExcelAttributes> mockedAttrs = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttrs.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(newAttrs);

            String result = resource.saveCredentials(payload);

            AuthenticationResponse response = Constants.GSON.fromJson(result, AuthenticationResponse.class);
            assertTrue(response.isSuccess());
            verify(refreshTokenStore, never()).remove(anyString());
        }
    }

    // ==================== saveCredentials — non-S2S MustAuthorizeException re-throws ====================

    @Test
    void saveCredentials_nonS2sMustAuthorizeException_rethrows() throws IOException {
        JsonObject payload = new JsonObject();
        payload.addProperty(Constants.AUTH_TYPE, Constants.PUBLIC);
        payload.addProperty(Constants.EMAIL, "user@example.com");

        OfficeExcelAttributes attributes = mock(OfficeExcelAttributes.class);
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);

        when(attributes.getAuthFlow()).thenReturn(Constants.AUTH_FLOW_PUBLIC); // non-S2S → re-throw
        when(providerFactory.create(attributes)).thenReturn(clientProvider);
        when(clientProvider.getGraphServiceClientForAdmin())
                .thenThrow(new MustAuthorizeException("Re-auth needed", Collections.emptyList()));
        when(officeExcelAttributeStore.save(attributes)).thenReturn("ctx-rethrow");

        try (MockedStatic<OfficeExcelAttributes> mockedAttrs = mockStatic(OfficeExcelAttributes.class)) {
            mockedAttrs.when(() -> OfficeExcelAttributes.create(payload, BASE_URL)).thenReturn(attributes);

            assertThrows(MustAuthorizeException.class, () -> resource.saveCredentials(payload));
            verify(officeExcelAttributeStore).remove("ctx-rethrow");
        }
    }

    // ==================== OAuth callback — key containing underscore ====================

    @Test
    void getCallBackForV2Auth_keyWithUnderscore_coversBothUnderscoreBranches() throws Exception {
        // state = "prefix_version#auth123" → key = "prefix_version" (contains _)
        // Line 141: userEmail = "prefix" (metadata triggeredBy)
        // Line 172: configuredEmail = "prefix" (email comparison)
        // JWT preferred_username = "prefix" → matches → authentication succeeds
        String state = "prefix_version#auth123";
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        TokenResult tokenResult = mock(TokenResult.class);

        when(officeExcelAttributeStore.load("auth123")).thenReturn(attrs);
        when(providerFactory.create(attrs)).thenReturn(clientProvider);
        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any())).thenReturn(tokenResult);

        // JWT with preferred_username = "prefix" — matches configuredEmail from split
        // Also: {"preferred_username":"prefix"} = 31 bytes → base64 = 42 chars → 42 % 4 = 2 → mod > 0 ✓
        String jwtPayload = "{\"preferred_username\":\"prefix\"}";
        String encodedPayload = java.util.Base64.getUrlEncoder().withoutPadding()
                .encodeToString(jwtPayload.getBytes());
        when(tokenResult.getAccessToken()).thenReturn("header." + encodedPayload + ".sig");
        when(authorizationContext.isAuthenticated()).thenReturn(false);

        String result = resource.getCallBackForV2Auth("code", state, null);

        assertTrue(result.contains(Constants.USER_AUTHENTICATED_SUCCESSFULLY));
        verify(officeExcelAttributeStore).remove("auth123");
    }

    // ==================== Helper Methods ====================

    /**
     * Sets up a successful Microsoft Auth SDK OAuth callback scenario.
     *
     * @param state           the full OAuth state string
     * @param authContextId   the auth context ID parsed from state (null if not present)
     * @param isAuthenticated whether the authorization context reports authenticated
     */
    private void setupSuccessfulOAuthCallback(String state, String authContextId, boolean isAuthenticated)
            throws Exception {
        GraphServiceClientProvider clientProvider = mock(GraphServiceClientProvider.class);
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk sdk = mock(MicrosoftAuthSdk.class);
        TokenResult tokenResult = mock(TokenResult.class);

        if (authContextId != null) {
            when(officeExcelAttributeStore.load(authContextId)).thenReturn(attrs);
            when(providerFactory.create(attrs)).thenReturn(clientProvider);
        } else {
            when(providerFactory.create()).thenReturn(clientProvider);
        }

        when(clientProvider.getOfficeExcelAttributes()).thenReturn(attrs);
        when(attrs.toSdk(any(ITokenStore.class))).thenReturn(sdk);
        when(sdk.handleCallback(eq("code"), any())).thenReturn(tokenResult);
        // For non-wsContact keys, email check now uses extractEmailFromAccessToken(tokenResult.getAccessToken()).
        // Returning null (Mockito default) causes extractEmailFromAccessToken to return null → check is skipped → allowed.
        when(authorizationContext.isAuthenticated()).thenReturn(isAuthenticated);
    }
}
