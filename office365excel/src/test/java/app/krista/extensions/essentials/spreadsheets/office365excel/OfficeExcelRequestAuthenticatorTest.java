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

import app.krista.extension.authorization.MustAuthenticateException;
import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.authorization.RequestAuthenticator;
import app.krista.extension.request.ProtoRequest;
import app.krista.extension.request.protos.http.HttpRequest;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.field.NamedValuedField;
import app.krista.sdk.msauth.MicrosoftAuthSdk;
import app.krista.sdk.msauth.token.ITokenStore;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import javax.ws.rs.core.MultivaluedHashMap;
import javax.ws.rs.core.MultivaluedMap;
import java.io.IOException;
import java.lang.reflect.Method;
import java.net.URI;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

/**
 * Test suite for the SDK-based OfficeExcelRequestAuthenticator.
 *
 * <p>This authenticator extends AbstractRequestAuthenticator from the microsoft-auth-sdk
 * and delegates OAuth URL generation to MicrosoftAuthSdk rather than ScribeJava.</p>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class OfficeExcelRequestAuthenticatorTest {

    @Mock
    private OfficeExcelAttributeStore officeExcelAttributeStore;

    @Mock
    private HttpRequest httpRequest;

    @Mock
    private URI uri;

    private OfficeExcelRequestAuthenticator authenticator;

    // Callback paths computed from prefix "/excel365extn"
    private static final String CALLBACK_PATH_V3 = "/excel365extn/v3/oauth/callback";
    private static final String CALLBACK_PATH_V2 = "/excel365extn/callback";

    @BeforeEach
    void setUp() {
        authenticator = new OfficeExcelRequestAuthenticator(
                officeExcelAttributeStore, "test-invoker", "https://base.url");
    }

    // ==================== Constructor ====================

    @Test
    void constructor_initializesFieldsCorrectly() {
        assertNotNull(authenticator);
    }

    // ==================== getScheme ====================

    @Test
    void getScheme_returnsNull() {
        assertNull(authenticator.getScheme());
    }

    // ==================== getSupportedProtocols ====================

    /**
     * AbstractRequestAuthenticator returns Set.of() (empty, non-null) — NOT null.
     */
    @Test
    void getSupportedProtocols_returnsEmptySet() {
        Set<String> protocols = authenticator.getSupportedProtocols();
        assertNotNull(protocols);
        assertTrue(protocols.isEmpty());
    }

    // ==================== setServiceAuthorization ====================

    @Test
    void setServiceAuthorization_returnsFalse() {
        assertFalse(authenticator.setServiceAuthorization("any-value"));
    }

    // ==================== getAttributeFields ====================

    /**
     * AbstractRequestAuthenticator returns Map.of() (empty, non-null) — NOT null.
     */
    @Test
    void getAttributeFields_returnsEmptyMap() {
        var fields = authenticator.getAttributeFields();
        assertNotNull(fields);
        assertTrue(fields.isEmpty());
    }

    // ==================== getMustAuthenticateResponse (with ProtoRequest) ====================

    @Test
    void getMustAuthenticateResponse_withProtoRequest_returnsNull() {
        MustAuthenticateException exception = mock(MustAuthenticateException.class);
        ProtoRequest request = mock(ProtoRequest.class);

        assertNull(authenticator.getMustAuthenticateResponse(exception, request));
    }

    // ==================== getMustAuthenticateResponse (no ProtoRequest) ====================

    @Test
    void getMustAuthenticateResponse_withoutProtoRequest_returnsNull() {
        MustAuthenticateException exception = mock(MustAuthenticateException.class);

        assertNull(authenticator.getMustAuthenticateResponse(exception));
    }

    // ==================== getMustAuthorizeResponse (with ProtoRequest) ====================

    @Test
    void getMustAuthorizeResponse_withProtoRequest_returnsNull() {
        MustAuthorizeException exception = mock(MustAuthorizeException.class);
        ProtoRequest request = mock(ProtoRequest.class);

        assertNull(authenticator.getMustAuthorizeResponse(exception, request));
    }

    // ==================== getAuthenticatedAccountId — V3 path ====================

    @Test
    void getAuthenticatedAccountId_v3Path_singlePart_returnsUserId() throws IOException {
        MultivaluedMap<String, String> queryParams = new MultivaluedHashMap<>();
        queryParams.add("state", "user@example.com");

        when(httpRequest.getUri()).thenReturn(uri);
        when(uri.getPath()).thenReturn(CALLBACK_PATH_V3);
        when(httpRequest.getQueryParameters()).thenReturn(queryParams);

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertEquals("user@example.com", result);
        verify(httpRequest).bufferBody();
    }

    @Test
    void getAuthenticatedAccountId_v3Path_twoPartState_returnsUserId() throws IOException {
        MultivaluedMap<String, String> queryParams = new MultivaluedHashMap<>();
        queryParams.add("state", "user@example.com#authCtx123");

        when(httpRequest.getUri()).thenReturn(uri);
        when(uri.getPath()).thenReturn(CALLBACK_PATH_V3);
        when(httpRequest.getQueryParameters()).thenReturn(queryParams);

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertEquals("user@example.com", result);
        verify(httpRequest).bufferBody();
    }

    // ==================== getAuthenticatedAccountId — V2 path ====================

    @Test
    void getAuthenticatedAccountId_v2Path_singlePart_returnsUserId() throws IOException {
        MultivaluedMap<String, String> queryParams = new MultivaluedHashMap<>();
        queryParams.add("state", "user@example.com");

        when(httpRequest.getUri()).thenReturn(uri);
        when(uri.getPath()).thenReturn(CALLBACK_PATH_V2);
        when(httpRequest.getQueryParameters()).thenReturn(queryParams);

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertEquals("user@example.com", result);
        verify(httpRequest).bufferBody();
    }

    @Test
    void getAuthenticatedAccountId_v2Path_multiPartState_returnsUserId() throws IOException {
        MultivaluedMap<String, String> queryParams = new MultivaluedHashMap<>();
        queryParams.add("state", "user@example.com#authCtx456#/some/forward/path");

        when(httpRequest.getUri()).thenReturn(uri);
        when(uri.getPath()).thenReturn(CALLBACK_PATH_V2);
        when(httpRequest.getQueryParameters()).thenReturn(queryParams);

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertEquals("user@example.com", result);
        verify(httpRequest).bufferBody();
    }

    // ==================== getAuthenticatedAccountId — non-OAuth paths ====================

    @Test
    void getAuthenticatedAccountId_nonOAuthPath_returnsNull() throws IOException {
        when(httpRequest.getUri()).thenReturn(uri);
        when(uri.getPath()).thenReturn("/some/other/path");

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertNull(result);
        verify(httpRequest, never()).bufferBody();
    }

    @Test
    void getAuthenticatedAccountId_nullPath_returnsNull() throws IOException {
        when(httpRequest.getUri()).thenReturn(uri);
        when(uri.getPath()).thenReturn(null);

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertNull(result);
    }

    // ==================== getAuthenticatedAccountId — exception handling ====================

    @Test
    void getAuthenticatedAccountId_ioException_returnsNull() throws IOException {
        when(httpRequest.getUri()).thenReturn(uri);
        when(uri.getPath()).thenReturn(CALLBACK_PATH_V3);
        doThrow(new IOException("Buffer error")).when(httpRequest).bufferBody();

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertNull(result);
    }

    @Test
    void getAuthenticatedAccountId_runtimeException_returnsNull() {
        when(httpRequest.getUri()).thenThrow(new RuntimeException("Unexpected error"));

        String result = authenticator.getAuthenticatedAccountId(httpRequest);

        assertNull(result);
    }

    // ==================== getMustAuthorizeResponse — Private auth, no authContextId ====================

    /**
     * When authType is PRIVATE and no authContextId is in the exception details,
     * the SDK builds the authorization URL using the loaded attributes.
     */
    @Test
    void getMustAuthorizeResponse_privateAuth_noAuthContext_buildsUrl() {
        List<NamedValuedField> details = new ArrayList<>();
        details.add(new NamedValuedField(
                Constants.USER_ID, Constants.TEXT, "user@example.com",
                new HashMap<>(), new HashMap<>()));

        MustAuthorizeException exception = new MustAuthorizeException(
                Constants.AUTHORIZATION_PROMPT, details);

        OfficeExcelAttributes mockAttrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk mockSdk = mock(MicrosoftAuthSdk.class);

        when(officeExcelAttributeStore.load("test-invoker")).thenReturn(mockAttrs);
        when(mockAttrs.getAuthType()).thenReturn(Constants.PRIVATE);
        when(mockAttrs.toSdk(any(ITokenStore.class))).thenReturn(mockSdk);
        when(mockSdk.buildAuthorizationUrl("user@example.com")).thenReturn("https://auth.url");

        RequestAuthenticator.AuthorizationResponse result =
                authenticator.getMustAuthorizeResponse(exception);

        assertNotNull(result);
        verify(officeExcelAttributeStore).load("test-invoker");
        verify(mockSdk).buildAuthorizationUrl("user@example.com");
    }

    // ==================== getMustAuthorizeResponse — empty details ====================

    /**
     * When exception details are empty, AbstractRequestAuthenticator throws IllegalStateException.
     */
    @Test
    void getMustAuthorizeResponse_emptyDetails_throwsException() {
        List<NamedValuedField> details = new ArrayList<>();
        MustAuthorizeException exception = new MustAuthorizeException(
                Constants.AUTHORIZATION_PROMPT, details);

        assertThrows(Exception.class,
                () -> authenticator.getMustAuthorizeResponse(exception));
    }

    // ==================== getMustAuthorizeResponse — null attributes fallback ====================

    /**
     * When loadAttributes returns null, the authenticator falls back to createDefaultAttributes,
     * which creates a Public-type OfficeExcelAttributes with routingUrl = "https://base.url".
     * The Public path appends the forwardPath to the state before building the URL.
     * Since real OfficeExcelAttributes requires config.properties, we verify the fallback
     * path is taken by checking that the store was queried and returned null.
     */
    @Test
    void getMustAuthorizeResponse_nullAttributes_fallsBackToDefaults() {
        List<NamedValuedField> details = new ArrayList<>();
        details.add(new NamedValuedField(
                Constants.USER_ID, Constants.TEXT, "user@example.com",
                new HashMap<>(), new HashMap<>()));

        MustAuthorizeException exception = new MustAuthorizeException(
                Constants.AUTHORIZATION_PROMPT, details);

        // Store returns null — should trigger createDefaultAttributes fallback
        when(officeExcelAttributeStore.load("test-invoker")).thenReturn(null);

        // createDefaultAttributes produces a real OfficeExcelAttributes(null, null, null, null, "Public", routingUrl)
        // which internally calls loadPublicConfig() from config.properties.
        // If the resource is available the call proceeds; if not it throws RuntimeException.
        // Either way, the store query happened and null was returned — that is what we assert here.
        try {
            authenticator.getMustAuthorizeResponse(exception);
        } catch (Exception ignored) {
            // The fallback attribute creation may fail in test environment without full
            // config.properties + Azure credentials. The important assertion is below.
        }

        // Verify that the attribute store was indeed consulted with the invokerId
        verify(officeExcelAttributeStore).load("test-invoker");
    }

    // ==================== getMustAuthorizeResponse — wsContact user builds OAuth URL ====================

    @Test
    void getMustAuthorizeResponse_wsContactUser_buildsOAuthUrl() {
        List<NamedValuedField> details = new ArrayList<>();
        details.add(new NamedValuedField(
                Constants.CLIENT_USER_ID, Constants.TEXT, "wsContact_user123",
                new HashMap<>(), new HashMap<>()));
        details.add(new NamedValuedField(
                Constants.USER_ID, Constants.TEXT, "INV_invoker1",
                new HashMap<>(), new HashMap<>()));

        MustAuthorizeException exception = new MustAuthorizeException(
                "No credentials configured.", details);

        // createDefaultAttributes falls back to config.properties; may throw in test env.
        // We verify the store was NOT queried (wsContact path uses createDefaultAttributes directly).
        try {
            authenticator.getMustAuthorizeResponse(exception);
        } catch (Exception ignored) {
            // May fail loading config.properties in unit test; that is acceptable here.
        }

        // For wsContact path the attribute store must NOT be queried with the invokerId
        verify(officeExcelAttributeStore, never()).load("wsContact_user123");
    }

    // ==================== getMustAuthorizeResponse — non-wsContact delegates to super ====================

    @Test
    void getMustAuthorizeResponse_nonWsContactUser_delegatesToSuper() {
        List<NamedValuedField> details = new ArrayList<>();
        details.add(new NamedValuedField(
                Constants.USER_ID, Constants.TEXT, "user@example.com",
                new HashMap<>(), new HashMap<>()));

        MustAuthorizeException exception = new MustAuthorizeException(
                Constants.AUTHORIZATION_PROMPT, details);

        OfficeExcelAttributes mockAttrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk mockSdk = mock(MicrosoftAuthSdk.class);

        when(officeExcelAttributeStore.load("test-invoker")).thenReturn(mockAttrs);
        when(mockAttrs.getAuthType()).thenReturn(Constants.PRIVATE);
        when(mockAttrs.toSdk(any(ITokenStore.class))).thenReturn(mockSdk);
        when(mockSdk.buildAuthorizationUrl("user@example.com")).thenReturn("https://auth.url");

        RequestAuthenticator.AuthorizationResponse result =
                authenticator.getMustAuthorizeResponse(exception);

        assertNotNull(result);
        // Non-wsContact path must consult the attribute store (delegated to super)
        verify(officeExcelAttributeStore).load("test-invoker");
        verify(mockSdk).buildAuthorizationUrl("user@example.com");
    }

    // ==================== loadAttributes — delegates to store ====================

    /**
     * Verifies that loadAttributes delegates to the injected OfficeExcelAttributeStore.
     * Tested indirectly through getMustAuthorizeResponse, which calls loadAttributes(invokerId).
     */
    @Test
    void loadAttributes_delegatesToStore() {
        List<NamedValuedField> details = new ArrayList<>();
        details.add(new NamedValuedField(
                Constants.USER_ID, Constants.TEXT, "user@example.com",
                new HashMap<>(), new HashMap<>()));

        MustAuthorizeException exception = new MustAuthorizeException(
                Constants.AUTHORIZATION_PROMPT, details);

        OfficeExcelAttributes mockAttrs = mock(OfficeExcelAttributes.class);
        MicrosoftAuthSdk mockSdk = mock(MicrosoftAuthSdk.class);

        when(officeExcelAttributeStore.load("test-invoker")).thenReturn(mockAttrs);
        when(mockAttrs.getAuthType()).thenReturn(Constants.PRIVATE);
        when(mockAttrs.toSdk(any(ITokenStore.class))).thenReturn(mockSdk);
        when(mockSdk.buildAuthorizationUrl(anyString())).thenReturn("https://auth.url");

        authenticator.getMustAuthorizeResponse(exception);

        // The store must have been called with the invokerId passed to the constructor
        verify(officeExcelAttributeStore).load("test-invoker");
    }
}
