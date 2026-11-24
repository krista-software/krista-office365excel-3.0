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

import app.krista.sdk.msauth.extension.base.MsAuthAttributes;
import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;

/**
 * Configuration container for Office365 Excel authentication attributes.
 * Implements MsAuthAttributes for Microsoft Auth SDK integration.
 */
public class OfficeExcelAttributes implements MsAuthAttributes {

    @SerializedName(CLIENT_ID)
    private String clientId;
    @SerializedName(CLIENT_SECRET)
    private String clientSecret;
    @SerializedName(TENANT_ID)
    private String tenantId;
    @SerializedName(EMAIL)
    private String emailId;
    @SerializedName(AUTH_TYPE)
    private String authType;
    @SerializedName(AUTH_FLOW)
    private String authFlow;
    @SerializedName("publicAuthMode")
    private String publicAuthMode;

    private String routingUrl;
    private String publicClientId;
    private String publicClientSecret;

    /**
     * Creates OfficeExcelAttributes from a JSON authentication payload sent by the SDK auth tab.
     *
     * <p>Handles three flows distinguished by {@code authFlow}:
     * <ul>
     *   <li>{@code AUTH_CODE_PUBLIC} — Quick Setup: Krista-managed credentials (clientId/Secret loaded from config)</li>
     *   <li>{@code AUTH_CODE_PRIVATE} — User Authentication: the user's own Azure AD app credentials</li>
     *   <li>{@code CLIENT_CREDENTIALS_SECRET} — Server to Server: app-only client credentials</li>
     * </ul>
     */
    public static OfficeExcelAttributes create(JsonObject authPayload, String baseurl) {
        String authType = authPayload.get(AUTH_TYPE).getAsString();
        String authFlow = authPayload.has(AUTH_FLOW) ? authPayload.get(AUTH_FLOW).getAsString() : null;
        String email = authPayload.has(EMAIL) ? authPayload.get(EMAIL).getAsString() : null;

        if (PUBLIC.equals(authType) && AUTH_FLOW_PRIVATE_USER.equals(authFlow)) {
            // User Authentication — own Azure AD app, auth code with user sign-in
            String clientId = authPayload.has(CLIENT_ID) ? authPayload.get(CLIENT_ID).getAsString() : null;
            String clientSecret = authPayload.has(CLIENT_SECRET) ? authPayload.get(CLIENT_SECRET).getAsString() : null;
            String tenantId = authPayload.has(TENANT_ID) ? authPayload.get(TENANT_ID).getAsString() : null;
            return new OfficeExcelAttributes(clientId, clientSecret, tenantId, email, authType, AUTH_FLOW_PRIVATE_USER, baseurl);
        } else if (PUBLIC.equals(authType)) {
            // Quick Setup — Krista-managed credentials (null triggers fallback to config.properties)
            return new OfficeExcelAttributes(null, null, null, email, authType, AUTH_FLOW_PUBLIC, baseurl);
        } else if (PRIVATE.equals(authType)) {
            // Server to Server — client credentials flow
            return new OfficeExcelAttributes(authPayload.get(CLIENT_ID).getAsString(),
                    authPayload.get(CLIENT_SECRET).getAsString(),
                    authPayload.get(TENANT_ID).getAsString(),
                    email, authType, AUTH_FLOW_CLIENT_CREDS, baseurl);
        } else {
            throw new IllegalArgumentException(UNSUPPORTED_AUTH);
        }
    }

    public static OfficeExcelAttributes create(String routingId, Map<String, Object> attributes) {
        String clientId = getStringOrNull(attributes, CLIENT_ID);
        String clientSecret = getStringOrNull(attributes, CLIENT_SECRET);
        String tenantId = getStringOrNull(attributes, TENANT_ID);
        String email = getStringOrNull(attributes, EMAIL);
        String authType = getStringOrNull(attributes, AUTH_TYPE);
        String authFlow = getStringOrNull(attributes, AUTH_FLOW);

        if (PUBLIC.equals(authType)) {
            // Backward compat: PUBLIC + own credentials (no explicit authFlow) = User Authentication.
            // PUBLIC + no own credentials = Quick Setup (Krista-managed).
            if (authFlow == null) {
                authFlow = (clientId != null && clientSecret != null) ? AUTH_FLOW_PRIVATE_USER : AUTH_FLOW_PUBLIC;
            }
            return new OfficeExcelAttributes(clientId, clientSecret, tenantId, email, authType, authFlow, routingId);
        }

        if (PRIVATE.equals(authType) && authFlow == null) {
            // Legacy migration: old Private tab (no authFlow stored) used auth code flow with own Azure AD app.
            // This is equivalent to the new User Authentication (AUTH_CODE_PRIVATE) flow.
            // Migrate authType to PUBLIC so getCredentials() matching works with the new SDK auth tab.
            return new OfficeExcelAttributes(clientId, clientSecret, tenantId, email, PUBLIC, AUTH_FLOW_PRIVATE_USER, routingId);
        }
        // Explicit S2S (authFlow=CLIENT_CREDENTIALS_SECRET), or legacy with no authType — preserve as S2S.
        return new OfficeExcelAttributes(clientId, clientSecret, tenantId, email, PRIVATE,
                authFlow != null ? authFlow : AUTH_FLOW_CLIENT_CREDS, routingId);
    }

    /**
     * Primary constructor — accepts an explicit authFlow string.
     */
    public OfficeExcelAttributes(String clientId, String clientSecret, String tenantId, String emailId,
                                 String authType, String authFlow, String routingUrl) {
        loadPublicConfig();
        this.clientId = clientId == null ? publicClientId : clientId;
        this.clientSecret = clientSecret == null ? publicClientSecret : clientSecret;
        this.tenantId = tenantId;
        this.emailId = emailId;
        this.authType = authType;
        this.authFlow = authFlow;
        this.publicAuthMode = AUTH_FLOW_PRIVATE_USER.equals(authFlow) ? "own"
                : AUTH_FLOW_PUBLIC.equals(authFlow) ? "krista"
                : null;
        this.routingUrl = routingUrl != null ? routingUrl.replace(LOCAL_EXTN_URL, LOCAL_EXTN_REPLACE_URL) : routingUrl;
    }

    /**
     * Backward-compatible constructor — derives authFlow from authType.
     */
    public OfficeExcelAttributes(String clientId, String clientSecret, String tenantId, String emailId,
                                 String authType, String routingUrl) {
        this(clientId, clientSecret, tenantId, emailId, authType,
                PUBLIC.equals(authType) ? AUTH_FLOW_PUBLIC : AUTH_FLOW_CLIENT_CREDS, routingUrl);
    }

    /**
     * Loads default public client configuration from properties file.
     */
    private void loadPublicConfig() {
        Properties properties = new Properties();
        try (InputStream inputStream = getClass().getClassLoader().getResourceAsStream("config.properties")) {
            properties.load(inputStream);
            this.publicClientId = properties.getProperty(PUBLIC_CLIENT_ID);
            this.publicClientSecret = properties.getProperty(PUBLIC_CLIENT_SECRET);
        } catch (IOException cause) {
            throw new RuntimeException("Failed to load public configuration file.", cause);
        }
    }

    @Override
    public String getScopeString() {
        return SCOPE;
    }

    @Override
    public String getCallbackEndPoint() {
        if (AUTH_FLOW_PRIVATE_USER.equals(authFlow)) {
            // User Authentication — redirect URI that users register in their own Azure AD app.
            // Uses the shorter /rest/excel365extn/callback path (same handler as v3/oauth/callback).
            return routingUrl + REST_EXCEL_CALLBACK;
        } else if (PUBLIC.equals(authType)) {
            // Quick Setup — centralized Krista platform callback
            return routingUrl.replaceAll(EXTENSION_API_PATTERN, DEFAULT_CALLBACK_PATH);
        } else {
            // Server to Server — no callback needed for client credentials flow
            return EMPTY_STRING;
        }
    }

    @Override
    public String getForwardPath() {
        if (AUTH_FLOW_PUBLIC.equals(authFlow)) {
            // Quick Setup uses a platform-level callback that forwards back to this extension path
            return routingUrl + EXTENSION_FORWARD_PATH;
        } else {
            return EMPTY_STRING;
        }
    }

    public String getClientId() {
        return clientId;
    }

    public String getClientSecret() {
        return clientSecret;
    }

    public String getTenantId() {
        return tenantId;
    }

    public String getEmailId() {
        return emailId;
    }

    public String getEmail() {
        return emailId;
    }

    public String getAuthType() {
        return authType;
    }

    public String getAuthFlow() {
        return authFlow;
    }

    public String getPublicAuthMode() {
        return publicAuthMode;
    }

    public void update(OfficeExcelAttributes officeExcelAttributes) {
        this.clientId = officeExcelAttributes.getClientId();
        this.clientSecret = officeExcelAttributes.getClientSecret();
        this.tenantId = officeExcelAttributes.getTenantId();
        this.emailId = officeExcelAttributes.getEmailId();
        this.authType = officeExcelAttributes.getAuthType();
    }

    public void update(Map<String, Object> newAttributes) {
        clientId = getStringOrNull(newAttributes, CLIENT_ID);
        clientSecret = getStringOrNull(newAttributes, CLIENT_SECRET);
        tenantId = getStringOrNull(newAttributes, TENANT_ID);
        emailId = getStringOrNull(newAttributes, EMAIL);
    }

    public Map<String, Object> toMap() {
        Map<String, Object> attributesMap = new HashMap<>();
        addAttributeIfNotNull(attributesMap, AUTH_TYPE, getAuthType());
        addAttributeIfNotNull(attributesMap, AUTH_FLOW, authFlow);
        addAttributeIfNotNull(attributesMap, "publicAuthMode", publicAuthMode);
        addAttributeIfNotNull(attributesMap, EMAIL, getEmail());
        // Include credentials only for flows where the user supplied their own Azure AD app.
        // Explicit positive check avoids the null-authFlow edge case where
        // !AUTH_FLOW_PUBLIC.equals(null) would incorrectly include Krista-managed credentials.
        boolean includeCredentials = AUTH_FLOW_PRIVATE_USER.equals(authFlow)
                || AUTH_FLOW_CLIENT_CREDS.equals(authFlow);
        if (includeCredentials) {
            addAttributeIfNotNull(attributesMap, CLIENT_ID, getClientId());
            addAttributeIfNotNull(attributesMap, CLIENT_SECRET, getClientSecret());
            addAttributeIfNotNull(attributesMap, TENANT_ID, getTenantId());
        }
        return attributesMap;
    }

    private static void addAttributeIfNotNull(Map<String, Object> map, String key, Object value) {
        if (value != null) {
            map.put(key, value);
        }
    }

    private static String getStringOrNull(Map<String, Object> map, String key) {
        Object val = map.get(key);
        return val instanceof String ? (String) val : null;
    }
}
