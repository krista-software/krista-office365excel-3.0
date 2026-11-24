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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores;

import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.util.KeyValueStore;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Tests for {@link OfficeExcelAttributeStore#load(String)}.
 *
 * Key regression: before the fix, loading attributes with authType=Public always discarded
 * clientId/clientSecret/tenantId and forced authFlow=AUTH_CODE_PUBLIC, causing User Authentication
 * (own Azure AD app, authFlow=AUTH_CODE_PRIVATE) to silently fall back to Krista's shared app.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class OfficeExcelAttributeStoreTest {

    @Mock private Invoker invoker;
    @Mock private RoutingInfo routingInfo;
    @Mock private KeyValueStore store;

    private OfficeExcelAttributeStore attributeStore;

    private static final String KEY = "INV_test-invoker";
    private static final String BASE_URL = "https://test.krista.app";

    @BeforeEach
    void setUp() {
        when(invoker.getRoutingInfo()).thenReturn(routingInfo);
        when(routingInfo.getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE))
                .thenReturn(BASE_URL);
        attributeStore = new OfficeExcelAttributeStore(invoker, store);
    }

    // ==================== Null / Empty value ====================

    @Test
    @DisplayName("load() returns null when store has no entry for key")
    void load_nullStoreValue_returnsNull() {
        when(store.get(KEY)).thenReturn(null);
        assertNull(attributeStore.load(KEY));
    }

    @Test
    @DisplayName("load() returns null when stored value is the string 'null'")
    void load_stringNullValue_returnsNull() {
        when(store.get(KEY)).thenReturn("null");
        assertNull(attributeStore.load(KEY));
    }

    @Test
    @DisplayName("load() returns null when stored value is blank")
    void load_blankValue_returnsNull() {
        when(store.get(KEY)).thenReturn("   ");
        assertNull(attributeStore.load(KEY));
    }

    // ==================== Quick Setup (Public + AUTH_CODE_PUBLIC) ====================

    @Test
    @DisplayName("load() Quick Setup: authFlow=AUTH_CODE_PUBLIC → null credentials (falls back to config)")
    void load_quickSetup_authFlowPublic_returnsPublicFlowWithNullUserCredentials() {
        String json = "{\"authType\":\"Public\",\"authFlow\":\"AUTH_CODE_PUBLIC\","
                + "\"email\":\"quick@example.com\"}";
        when(store.get(KEY)).thenReturn(json);

        OfficeExcelAttributes result = attributeStore.load(KEY);

        assertNotNull(result);
        assertEquals(PUBLIC, result.getAuthType());
        assertEquals(AUTH_FLOW_PUBLIC, result.getAuthFlow());
        assertEquals("quick@example.com", result.getEmail());
        // Quick Setup: the constructor falls back to Krista publicClientId from config.properties.
        // The important invariant is that no user-specific credentials are loaded for Quick Setup.
    }

    @Test
    @DisplayName("load() Quick Setup: no authFlow stored → defaults to AUTH_CODE_PUBLIC (backward compat)")
    void load_quickSetup_noAuthFlow_defaultsToAuthFlowPublic() {
        String json = "{\"authType\":\"Public\",\"email\":\"quick@example.com\"}";
        when(store.get(KEY)).thenReturn(json);

        OfficeExcelAttributes result = attributeStore.load(KEY);

        assertNotNull(result);
        assertEquals(PUBLIC, result.getAuthType());
        assertEquals(AUTH_FLOW_PUBLIC, result.getAuthFlow());
    }

    // ==================== User Authentication (Public + AUTH_CODE_PRIVATE) ====================
    // This is the key regression fix: before the fix, credentials were silently discarded here.

    @Test
    @DisplayName("load() User Auth: authFlow=AUTH_CODE_PRIVATE → PRESERVES user clientId (regression fix)")
    void load_userAuth_authFlowPrivate_preservesClientId() {
        String json = "{\"authType\":\"Public\",\"authFlow\":\"AUTH_CODE_PRIVATE\","
                + "\"email\":\"user@example.com\","
                + "\"clientId\":\"user-specific-client-id\","
                + "\"clientSecret\":\"user-specific-secret\","
                + "\"tenantId\":\"user-specific-tenant\"}";
        when(store.get(KEY)).thenReturn(json);

        OfficeExcelAttributes result = attributeStore.load(KEY);

        assertNotNull(result);
        assertEquals(PUBLIC, result.getAuthType());
        assertEquals(AUTH_FLOW_PRIVATE_USER, result.getAuthFlow());
        assertEquals("user@example.com", result.getEmail());
        // *** Regression check: these must be the user's own credentials, not Krista's publicClientId ***
        assertEquals("user-specific-client-id", result.getClientId());
        assertEquals("user-specific-secret", result.getClientSecret());
        assertEquals("user-specific-tenant", result.getTenantId());
    }

    @Test
    @DisplayName("load() User Auth: authFlow=AUTH_CODE_PRIVATE → authFlow preserved on returned attrs")
    void load_userAuth_authFlowPrivate_authFlowIsPrivateUser() {
        String json = "{\"authType\":\"Public\",\"authFlow\":\"AUTH_CODE_PRIVATE\","
                + "\"email\":\"user@example.com\","
                + "\"clientId\":\"cid\",\"clientSecret\":\"cs\",\"tenantId\":\"tid\"}";
        when(store.get(KEY)).thenReturn(json);

        OfficeExcelAttributes result = attributeStore.load(KEY);

        // authFlow must NOT be coerced to AUTH_CODE_PUBLIC — that was the original bug
        assertEquals(AUTH_FLOW_PRIVATE_USER, result.getAuthFlow());
        assertNotEquals(AUTH_FLOW_PUBLIC, result.getAuthFlow());
    }

    @Test
    @DisplayName("load() User Auth vs Quick Setup: same authType, different authFlow → different credential behaviour")
    void load_publicAuthType_authFlowDeterminesCredentials() {
        String quickSetupJson = "{\"authType\":\"Public\",\"authFlow\":\"AUTH_CODE_PUBLIC\","
                + "\"email\":\"quick@example.com\"}";
        String userAuthJson = "{\"authType\":\"Public\",\"authFlow\":\"AUTH_CODE_PRIVATE\","
                + "\"email\":\"user@example.com\","
                + "\"clientId\":\"my-client-id\",\"clientSecret\":\"my-secret\",\"tenantId\":\"my-tenant\"}";

        when(store.get("quick-key")).thenReturn(quickSetupJson);
        when(store.get("user-key")).thenReturn(userAuthJson);

        OfficeExcelAttributes quickSetup = attributeStore.load("quick-key");
        OfficeExcelAttributes userAuth   = attributeStore.load("user-key");

        // Quick Setup must NOT leak user-specific credentials from a different invokerId
        assertEquals(AUTH_FLOW_PUBLIC, quickSetup.getAuthFlow());

        // User Auth must preserve the user's own credentials
        assertEquals(AUTH_FLOW_PRIVATE_USER, userAuth.getAuthFlow());
        assertEquals("my-client-id", userAuth.getClientId());
        assertEquals("my-tenant", userAuth.getTenantId());
    }

    // ==================== Server to Server (Private + CLIENT_CREDENTIALS_SECRET) ====================

    @Test
    @DisplayName("load() S2S: authType=Private, authFlow=CLIENT_CREDENTIALS_SECRET → preserves all credentials")
    void load_serverToServer_preservesCredentials() {
        String json = "{\"authType\":\"Private\",\"authFlow\":\"CLIENT_CREDENTIALS_SECRET\","
                + "\"email\":\"svc@example.com\","
                + "\"clientId\":\"s2s-client-id\","
                + "\"clientSecret\":\"s2s-secret\","
                + "\"tenantId\":\"s2s-tenant\"}";
        when(store.get(KEY)).thenReturn(json);

        OfficeExcelAttributes result = attributeStore.load(KEY);

        assertNotNull(result);
        assertEquals(PRIVATE, result.getAuthType());
        assertEquals(AUTH_FLOW_CLIENT_CREDS, result.getAuthFlow());
        assertEquals("svc@example.com", result.getEmail());
        assertEquals("s2s-client-id", result.getClientId());
        assertEquals("s2s-secret", result.getClientSecret());
        assertEquals("s2s-tenant", result.getTenantId());
    }

    @Test
    @DisplayName("load() legacy Private (no authFlow) → migrates to User Authentication (AUTH_CODE_PRIVATE)")
    void load_legacyPrivate_noAuthFlow_migratestoUserAuthentication() {
        // Old Private tab used auth code flow (user sign-in with own Azure AD app),
        // not client credentials. Without an explicit authFlow, migrate to User Authentication.
        String json = "{\"authType\":\"Private\","
                + "\"email\":\"user@example.com\","
                + "\"clientId\":\"legacy-client-id\","
                + "\"clientSecret\":\"legacy-secret\","
                + "\"tenantId\":\"legacy-tenant\"}";
        when(store.get(KEY)).thenReturn(json);

        OfficeExcelAttributes result = attributeStore.load(KEY);

        assertNotNull(result);
        // authType migrated to PUBLIC so getCredentials() matching works with the new SDK auth tab
        assertEquals(PUBLIC, result.getAuthType());
        assertEquals(AUTH_FLOW_PRIVATE_USER, result.getAuthFlow());
        assertEquals("user@example.com", result.getEmail());
        assertEquals("legacy-client-id", result.getClientId());
        assertEquals("legacy-secret", result.getClientSecret());
        assertEquals("legacy-tenant", result.getTenantId());
    }

    // ==================== save() ====================

    @Test
    @DisplayName("save(attrs) generates a unique key prefixed with AUTH_CONTEXT_ID")
    void save_attrs_returnsGeneratedKey() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes(
                "cid", "cs", "tid", "user@example.com", PRIVATE, AUTH_FLOW_CLIENT_CREDS, BASE_URL);

        String key = attributeStore.save(attrs);

        assertNotNull(key);
        assertTrue(key.startsWith(AUTH_CONTEXT_ID + "_"),
                "Key should start with AUTH_CONTEXT_ID prefix, got: " + key);
        verify(store).put(eq(key), anyString());
    }

    @Test
    @DisplayName("save(attrs) throws IllegalArgumentException for null attributes")
    void save_nullAttrs_throwsIllegalArgument() {
        assertThrows(IllegalArgumentException.class, () -> attributeStore.save(null));
    }

    @Test
    @DisplayName("save(attrs, key) stores under specified key and returns true")
    void save_attrsWithKey_storesUnderGivenKey() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes(
                "cid", "cs", "tid", "user@example.com", PRIVATE, AUTH_FLOW_CLIENT_CREDS, BASE_URL);
        String specificKey = "my-specific-key";

        boolean result = attributeStore.save(attrs, specificKey);

        assertTrue(result);
        verify(store).put(eq(specificKey), anyString());
    }

    @Test
    @DisplayName("save(attrs, key) throws IllegalArgumentException for null attributes")
    void save_attrsWithKey_nullAttrs_throwsIllegalArgument() {
        assertThrows(IllegalArgumentException.class, () -> attributeStore.save(null, "some-key"));
    }

    // ==================== remove() ====================

    @Test
    @DisplayName("remove() delegates to store.remove()")
    void remove_delegatesToStore() {
        attributeStore.remove("ctx-123");
        verify(store).remove("ctx-123");
    }
}
