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

import app.krista.sdk.msauth.token.InMemoryTokenStore;
import com.google.gson.JsonObject;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.HashMap;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;
import static org.junit.jupiter.api.Assertions.*;

@ExtendWith(MockitoExtension.class)
class OfficeExcelAttributesTest {

    private static final String BASE_URL = "https://extension.local.eng.krista.app";

    @Test
    void testCreateWithUnsupportedAuthTypeThrowsException() {
        JsonObject payload = new JsonObject();
        payload.addProperty(AUTH_TYPE, "INVALID");
        payload.addProperty(EMAIL, "test@example.com");

        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                () -> OfficeExcelAttributes.create(payload, BASE_URL));
        assertEquals(UNSUPPORTED_AUTH, ex.getMessage());
    }

    @Test
    void testCreateFromRoutingUrlAndAttributesPopulatesFields() {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(CLIENT_ID, "client-id");
        attributes.put(CLIENT_SECRET, "client-secret");
        attributes.put(TENANT_ID, "tenant-id");
        attributes.put(EMAIL, "user@example.com");

        OfficeExcelAttributes result = OfficeExcelAttributes.create(BASE_URL, attributes);

        assertEquals("client-id", result.getClientId());
        assertEquals("client-secret", result.getClientSecret());
        assertEquals("tenant-id", result.getTenantId());
        assertEquals("user@example.com", result.getEmailId());
        assertEquals(PRIVATE, result.getAuthType());
        assertEquals(AUTH_FLOW_CLIENT_CREDS, result.getAuthFlow());
        // Server-to-Server (Client Credentials) requires no redirect URI
        assertEquals(EMPTY_STRING, result.getCallbackEndPoint());
        assertEquals(EMPTY_STRING, result.getForwardPath());
    }

    @Test
    void testCreateFromRoutingIdAndAttributes() {
        Map<String, Object> attrs = new HashMap<>();
        attrs.put(CLIENT_ID, "client-id");
        attrs.put(CLIENT_SECRET, "client-secret");
        attrs.put(TENANT_ID, "tenant-id");
        attrs.put(EMAIL, "user@example.com");
        String routingId = BASE_URL + "/extension/api/rest/v3/oauth/callback";

        OfficeExcelAttributes result = OfficeExcelAttributes.create(routingId, attrs);

        assertEquals("client-id", result.getClientId());
        assertEquals("client-secret", result.getClientSecret());
        assertEquals("tenant-id", result.getTenantId());
        assertEquals("user@example.com", result.getEmailId());
        assertEquals(PRIVATE, result.getAuthType());
        assertEquals(AUTH_FLOW_CLIENT_CREDS, result.getAuthFlow());
        // Server-to-Server (Client Credentials) requires no redirect URI
        assertEquals(EMPTY_STRING, result.getCallbackEndPoint());
    }

    @Test
    void testCallbackAndForwardPathForPublic() {
        String routingId = BASE_URL + "/extension/api/rest/v3/oauth/callback";
        OfficeExcelAttributes attrs = new OfficeExcelAttributes(null, null, null,
                "user@example.com", PUBLIC, routingId);

        String callbackEndPoint = attrs.getCallbackEndPoint();
        String forwardPath = attrs.getForwardPath();

        assertEquals(LOCAL_EXTN_REPLACE_URL + "/extension/api/rest/v3/oauth/callback", callbackEndPoint);
        assertEquals(LOCAL_EXTN_REPLACE_URL + "/extension/api/rest/v3/oauth/callback" + EXTENSION_FORWARD_PATH,
                forwardPath);
    }

    @Test
    void testCallbackAndForwardPathForPrivate() {
        // PRIVATE authType = Server-to-Server (Client Credentials) — no redirect URI needed
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("id", "secret", "tenant",
                "user@example.com", PRIVATE, LOCAL_EXTN_URL);

        assertEquals(EMPTY_STRING, attrs.getCallbackEndPoint());
        assertEquals(EMPTY_STRING, attrs.getForwardPath());
    }

    @Test
    void testCallbackAndForwardPathForUserAuthentication() {
        // User Authentication (AUTH_CODE_PRIVATE) uses the shorter /rest/excel365extn/callback path
        // so users register this simpler URL in their own Azure AD app.
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("id", "secret", "tenant",
                "user@example.com", PUBLIC, AUTH_FLOW_PRIVATE_USER, LOCAL_EXTN_URL);

        assertEquals(LOCAL_EXTN_REPLACE_URL + REST_EXCEL_CALLBACK, attrs.getCallbackEndPoint());
        assertEquals(EMPTY_STRING, attrs.getForwardPath());
    }

    @Test
    void testCallbackAndForwardPathForUnknownAuthType() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("id", "secret", "tenant",
                "user@example.com", "Unknown", LOCAL_EXTN_URL);

        assertEquals(EMPTY_STRING, attrs.getCallbackEndPoint());
        assertEquals(EMPTY_STRING, attrs.getForwardPath());
    }

    @Test
    void testUpdateFromOtherAttributes() {
        OfficeExcelAttributes target = new OfficeExcelAttributes("id1", "secret1", "tenant1",
                "email1", PRIVATE, LOCAL_EXTN_URL);
        OfficeExcelAttributes source = new OfficeExcelAttributes("id2", "secret2", "tenant2",
                "email2", PUBLIC, LOCAL_EXTN_URL);

        target.update(source);

        assertEquals("id2", target.getClientId());
        assertEquals("secret2", target.getClientSecret());
        assertEquals("tenant2", target.getTenantId());
        assertEquals("email2", target.getEmailId());
        assertEquals(PUBLIC, target.getAuthType());
    }

    @Test
    void testUpdateFromMapAndToMapOmitNullValues() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("id1", "secret1", null,
                "email1", PRIVATE, LOCAL_EXTN_URL);

        Map<String, Object> newAttrs = new HashMap<>();
        newAttrs.put(CLIENT_ID, "id2");
        newAttrs.put(CLIENT_SECRET, "secret2");
        newAttrs.put(EMAIL, "email2");

        attrs.update(newAttrs);

        assertEquals("id2", attrs.getClientId());
        assertEquals("secret2", attrs.getClientSecret());
        assertNull(attrs.getTenantId());
        assertEquals("email2", attrs.getEmailId());

        Map<String, Object> map = attrs.toMap();
        assertEquals("id2", map.get(CLIENT_ID));
        assertEquals("secret2", map.get(CLIENT_SECRET));
        assertEquals("email2", map.get(EMAIL));
        assertEquals(PRIVATE, map.get(AUTH_TYPE));
        assertFalse(map.containsKey(TENANT_ID));
    }

    @Test
    void testToSdkReturnsNonNull() {
        JsonObject publicPayload = new JsonObject();
        publicPayload.addProperty(AUTH_TYPE, PUBLIC);
        publicPayload.addProperty(EMAIL, "user@example.com");
        OfficeExcelAttributes publicAttrs = OfficeExcelAttributes.create(publicPayload, BASE_URL);

        assertNotNull(publicAttrs.toSdk(new InMemoryTokenStore()));
    }

    // ==================== create() factory tests ====================

    @Test
    void testCreateFromPayload_quickSetup_setsAuthFlowPublicAndKristaMode() {
        JsonObject payload = new JsonObject();
        payload.addProperty(AUTH_TYPE, PUBLIC);
        payload.addProperty(EMAIL, "user@example.com");

        OfficeExcelAttributes attrs = OfficeExcelAttributes.create(payload, BASE_URL);

        assertEquals(PUBLIC, attrs.getAuthType());
        assertEquals(AUTH_FLOW_PUBLIC, attrs.getAuthFlow());
        assertEquals("krista", attrs.getPublicAuthMode());
        assertEquals("user@example.com", attrs.getEmailId());
        // Quick Setup uses Krista-managed credentials (loaded from config.properties)
        assertNotNull(attrs.getClientId());
        assertNotNull(attrs.getClientSecret());
    }

    @Test
    void testCreateFromPayload_userAuth_setsAuthFlowPrivateUserAndOwnMode() {
        JsonObject payload = new JsonObject();
        payload.addProperty(AUTH_TYPE, PUBLIC);
        payload.addProperty(AUTH_FLOW, AUTH_FLOW_PRIVATE_USER);
        payload.addProperty(EMAIL, "user@example.com");
        payload.addProperty(CLIENT_ID, "user-client-id");
        payload.addProperty(CLIENT_SECRET, "user-client-secret");
        payload.addProperty(TENANT_ID, "user-tenant-id");

        OfficeExcelAttributes attrs = OfficeExcelAttributes.create(payload, BASE_URL);

        assertEquals(PUBLIC, attrs.getAuthType());
        assertEquals(AUTH_FLOW_PRIVATE_USER, attrs.getAuthFlow());
        assertEquals("own", attrs.getPublicAuthMode());
        assertEquals("user-client-id", attrs.getClientId());
        assertEquals("user-client-secret", attrs.getClientSecret());
        assertEquals("user-tenant-id", attrs.getTenantId());
        assertEquals("user@example.com", attrs.getEmailId());
    }

    @Test
    void testCreateFromPayload_serverToServer_setsAuthFlowClientCredsAndNullMode() {
        JsonObject payload = new JsonObject();
        payload.addProperty(AUTH_TYPE, PRIVATE);
        payload.addProperty(CLIENT_ID, "app-client-id");
        payload.addProperty(CLIENT_SECRET, "app-client-secret");
        payload.addProperty(TENANT_ID, "app-tenant-id");
        payload.addProperty(EMAIL, "admin@example.com");

        OfficeExcelAttributes attrs = OfficeExcelAttributes.create(payload, BASE_URL);

        assertEquals(PRIVATE, attrs.getAuthType());
        assertEquals(AUTH_FLOW_CLIENT_CREDS, attrs.getAuthFlow());
        assertNull(attrs.getPublicAuthMode());
        assertEquals("app-client-id", attrs.getClientId());
        assertEquals("app-client-secret", attrs.getClientSecret());
        assertEquals("app-tenant-id", attrs.getTenantId());
    }

    // ==================== toMap() conditional credential tests ====================

    @Test
    void testToMap_quickSetup_excludesCredentials() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes(null, null, null,
                "user@example.com", PUBLIC, AUTH_FLOW_PUBLIC, BASE_URL);

        Map<String, Object> map = attrs.toMap();

        assertEquals(PUBLIC, map.get(AUTH_TYPE));
        assertEquals(AUTH_FLOW_PUBLIC, map.get(AUTH_FLOW));
        assertEquals("krista", map.get("publicAuthMode"));
        assertEquals("user@example.com", map.get(EMAIL));
        // Krista-managed credentials must NEVER appear in the map
        assertFalse(map.containsKey(CLIENT_ID), "Quick Setup toMap() must not include clientId");
        assertFalse(map.containsKey(CLIENT_SECRET), "Quick Setup toMap() must not include clientSecret");
        assertFalse(map.containsKey(TENANT_ID), "Quick Setup toMap() must not include tenantId");
    }

    @Test
    void testToMap_userAuth_includesCredentials() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("my-client-id", "my-secret", "my-tenant",
                "user@example.com", PUBLIC, AUTH_FLOW_PRIVATE_USER, BASE_URL);

        Map<String, Object> map = attrs.toMap();

        assertEquals(PUBLIC, map.get(AUTH_TYPE));
        assertEquals(AUTH_FLOW_PRIVATE_USER, map.get(AUTH_FLOW));
        assertEquals("own", map.get("publicAuthMode"));
        assertEquals("my-client-id", map.get(CLIENT_ID));
        assertEquals("my-secret", map.get(CLIENT_SECRET));
        assertEquals("my-tenant", map.get(TENANT_ID));
    }

    @Test
    void testToMap_serverToServer_includesCredentials() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("s2s-client", "s2s-secret", "s2s-tenant",
                "admin@example.com", PRIVATE, AUTH_FLOW_CLIENT_CREDS, BASE_URL);

        Map<String, Object> map = attrs.toMap();

        assertEquals(PRIVATE, map.get(AUTH_TYPE));
        assertEquals(AUTH_FLOW_CLIENT_CREDS, map.get(AUTH_FLOW));
        assertNull(map.get("publicAuthMode"), "Server-to-Server should not have publicAuthMode");
        assertEquals("s2s-client", map.get(CLIENT_ID));
        assertEquals("s2s-secret", map.get(CLIENT_SECRET));
        assertEquals("s2s-tenant", map.get(TENANT_ID));
    }

    @Test
    void testToMap_nullAuthFlow_excludesCredentials() {
        // Edge case: authFlow is null (legacy data) — credentials should NOT leak
        OfficeExcelAttributes attrs = new OfficeExcelAttributes(null, null, null,
                "user@example.com", PUBLIC, BASE_URL);
        // Backward-compat constructor derives authFlow=AUTH_FLOW_PUBLIC → publicAuthMode="krista"

        Map<String, Object> map = attrs.toMap();

        assertFalse(map.containsKey(CLIENT_ID), "Null authFlow path must not include clientId");
        assertFalse(map.containsKey(CLIENT_SECRET), "Null authFlow path must not include clientSecret");
    }

    @Test
    void testGetScopeString_returnsScopeConstant() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes(null, null, null,
                "user@example.com", PUBLIC, BASE_URL);

        assertEquals(SCOPE, attrs.getScopeString());
    }

    @Test
    void testToMap_returnsNewMapInstance() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("cid", "cs", "tid",
                "user@example.com", PRIVATE, AUTH_FLOW_CLIENT_CREDS, BASE_URL);

        Map<String, Object> map1 = attrs.toMap();
        Map<String, Object> map2 = attrs.toMap();

        assertNotSame(map1, map2);
        assertEquals(map1, map2);
    }

    @Test
    void testToMap_nullEmail_omitted() {
        OfficeExcelAttributes attrs = new OfficeExcelAttributes("cid", "cs", "tid",
                null, PRIVATE, AUTH_FLOW_CLIENT_CREDS, BASE_URL);

        Map<String, Object> map = attrs.toMap();

        assertFalse(map.containsKey(EMAIL), "Null email should be omitted");
    }

    // ==================== create() from routing with attributes ====================

    @Test
    void testCreateFromRoutingAndAttributes_withAuthFlow_preservesAuthFlow() {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(CLIENT_ID, "cid");
        attributes.put(CLIENT_SECRET, "cs");
        attributes.put(TENANT_ID, "tid");
        attributes.put(EMAIL, "user@example.com");
        attributes.put(AUTH_FLOW, AUTH_FLOW_PRIVATE_USER);

        OfficeExcelAttributes result = OfficeExcelAttributes.create(BASE_URL, attributes);

        assertEquals(AUTH_FLOW_PRIVATE_USER, result.getAuthFlow());
        assertEquals(PRIVATE, result.getAuthType());
    }

    @Test
    void testCreateFromRoutingAndAttributes_noAuthFlow_defaultsToClientCreds() {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(CLIENT_ID, "cid");
        attributes.put(CLIENT_SECRET, "cs");

        OfficeExcelAttributes result = OfficeExcelAttributes.create(BASE_URL, attributes);

        assertEquals(AUTH_FLOW_CLIENT_CREDS, result.getAuthFlow());
    }

    @Test
    void testCreateFromRoutingAndAttributes_nonStringValue_treatedAsNull() {
        Map<String, Object> attributes = new HashMap<>();
        attributes.put(CLIENT_ID, 12345); // non-String value

        OfficeExcelAttributes result = OfficeExcelAttributes.create(BASE_URL, attributes);

        // Non-String values should be treated as null by getStringOrNull()
        // and trigger fallback to publicClientId from config.properties
        assertNotNull(result.getClientId());
    }
}

