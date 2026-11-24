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
import app.krista.sdk.msauth.extension.registry.ExtensionRegistry;
import app.krista.sdk.msauth.token.ITokenStore;
import app.krista.sdk.msauth.token.TokenResult;
import com.google.gson.JsonObject;
import com.microsoft.graph.http.GraphServiceException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import javax.ws.rs.*;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Base64;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.MicrosoftAuthErrorConstants.PERMISSIONS_REVOKED_ERROR;

/**
 * REST resource for handling Office365 Excel authentication and OAuth callbacks.
 *
 * <p>Migrated from ScribeJava-based OAuthService to the Microsoft Auth SDK.
 * Auth tab HTML is served from the SDK JAR (no React/Node build needed).</p>
 */
@Path("/")
public final class AuthenticationResource {

    private static final Logger LOGGER = LoggerFactory.getLogger(AuthenticationResource.class);
    public static final String USER_AUTHENTICATED_SUCCESSFULLY_PLEASE_PROCEED_WITH_REQUEST =
            "User Authenticated Successfully. Please proceed with the request.";

    private final OfficeExcelAttributeStore officeExcelAttributeStore;
    private final RefreshTokenStore refreshTokenStore;
    private final GraphServiceClientProviderFactory providerFactory;
    private final AuthorizationContext context;
    private final AuthorizationListener authorizationListener;
    private final String invokerId;
    private final String baseRoutingUrl;
    private final Office365ExcelAuditService auditService;

    @Inject
    public AuthenticationResource(OfficeExcelAttributeStore officeExcelAttributeStore,
                                  RefreshTokenStore refreshTokenStore,
                                  GraphServiceClientProviderFactory providerFactory,
                                  AuthorizationContext context,
                                  AuthorizationListener authorizationListener,
                                  Invoker invoker,
                                  Office365ExcelAuditService auditService) {
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.refreshTokenStore = refreshTokenStore;
        this.context = context;
        this.authorizationListener = authorizationListener;
        this.providerFactory = providerFactory;
        this.invokerId = invoker.getInvokerId();
        this.baseRoutingUrl = invoker.getRoutingInfo().getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE);
        this.auditService = auditService;
    }

    // ── Legacy callback (v2 auth flow) ─────────────────────────────────────────

    @GET
    @Path("callback")
    public String getCallBackForV2Auth(@QueryParam("code") String code,
                                       @QueryParam("state") String state,
                                       @QueryParam("admin_consent") String adminConsent) {
        if ("True".equals(adminConsent)) {
            auditService.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", "OAuth Callback",
                    "Admin consent granted", null, Map.of("triggeredBy", "Admin"));
            return "Admin consent granted successfully! Click 'Test Connection' to verify, then 'Save' to apply.";
        }
        return getAuthenticationResponseMessage(code, state);
    }

    // ── v3 callback (current auth flow) ───────────────────────────────────────

    @GET
    @Path("v3/oauth/callback")
    public String getCallbackForV3Auth(@QueryParam("code") String code,
                                       @QueryParam("state") String state,
                                       @QueryParam("admin_consent") String adminConsent) {
        if ("True".equals(adminConsent)) {
            auditService.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", "OAuth Callback",
                    "Admin consent granted", null, Map.of("triggeredBy", "Admin"));
            return "Admin consent granted successfully! Click 'Test Connection' to verify, then 'Save' to apply.";
        }
        return getAuthenticationResponseMessage(code, state);
    }

    // ── Core OAuth callback handler ────────────────────────────────────────────

    private String getAuthenticationResponseMessage(String code, String state) {
        auditService.startRequest(invokerId, "OAuth Callback");
        LOGGER.info("Authentication callback triggered with Code: [{}], State: [{}]", code, state);

        if (code == null) {
            LOGGER.debug("Missing authorization code in callback. Rejecting authentication.");
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "OAuth Callback",
                    "No authorization code received — user may have cancelled", null,
                    Map.of("triggeredBy", "Unknown"));
            return "Authentication Failed. Please re-authorize.";
        }

        String[] parts = state != null ? state.split(HASH) : new String[]{""};
        if (parts[0].isBlank()) {
            LOGGER.error("State parameter is invalid or empty.");
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "OAuth Callback",
                    "Invalid state parameter — possible CSRF or corrupted redirect", null,
                    Map.of("triggeredBy", "Unknown"));
            throw new BadRequestException("Invalid OAuth state. Please restart the authentication flow.");
        }

        String key = parts[0];
        String clientKey = null;
        if (parts.length >= 3 && key.startsWith(WS_CONTACT)) {
            clientKey = parts[1]; // invokerId (e.g. INV_xxx)
        }
        // For wsContact, parts[1] is the invokerId — not an authContextId
        String authContextId = key.startsWith(WS_CONTACT) ? null : getAuthContextId(state);

        String userEmail = key.contains(UNDER_SCORE) ? key.split(UNDER_SCORE)[0] : key;
        Map<String, String> callbackMeta = new HashMap<>();
        callbackMeta.put("triggeredBy", userEmail);

        LOGGER.debug("Parsed key: [{}] from state. AuthContextId resolved: [{}]", key, authContextId);

        try {
            GraphServiceClientProvider clientProvider = getCallbackProvider(authContextId, key);
            OfficeExcelAttributes officeExcelAttributes = clientProvider.getOfficeExcelAttributes();

            LOGGER.debug("Exchanging authorization code via MicrosoftAuthSdk...");
            ITokenStore tokenStore = createTokenStore();
            MicrosoftAuthSdk sdk = officeExcelAttributes.toSdk(tokenStore);
            TokenResult tokenResult = sdk.handleCallback(code, key);

            // MCP path: save permanent Public attrs under wsContactId + suffix.
            // The suffix avoids key collision with the refresh token, which the SDK
            // stores under bare wsContactId in the same KeyValueStore.
            if (clientKey != null && key.startsWith(WS_CONTACT)) {
                String userEmailFromToken = extractEmailFromAccessToken(tokenResult.getAccessToken());
                String attrsKey = key + Constants.PER_USER_ATTRS_SUFFIX;
                OfficeExcelAttributes enrichedAttrs = new OfficeExcelAttributes(null, null, null, userEmailFromToken, PUBLIC, baseRoutingUrl);
                officeExcelAttributeStore.save(enrichedAttrs, attrsKey);
                LOGGER.info("MCP callback: saving per-user attrs with email={} under attrsKey={} (token under {})",
                        userEmailFromToken, attrsKey, key);
            }
            LOGGER.info("Access token retrieved successfully via SDK.");

            if (!key.startsWith(WS_CONTACT)) {
                String authenticatedEmail = extractEmailFromAccessToken(tokenResult.getAccessToken());
                if (authenticatedEmail != null && !authenticatedEmail.isBlank()) {
                    String configuredEmail = key.contains(UNDER_SCORE) ? key.split(UNDER_SCORE)[0] : key;
                    if (!authenticatedEmail.equalsIgnoreCase(configuredEmail)) {
                        LOGGER.info("User access denied. Authenticated email '{}' does not match configured email '{}'.",
                                authenticatedEmail, configuredEmail);
                        refreshTokenStore.remove(key);
                        auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "OAuth Callback",
                                "Sign-in rejected: user signed in but setup is configured for '" + configuredEmail + "'",
                                null, callbackMeta);
                        return UNAUTHORISED_USER + " User email '" + configuredEmail
                                + "' configured in the setup does not match with authenticated user.";
                    }
                }
            }

            auditService.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", "OAuth Callback",
                    "User authenticated: " + callbackMeta.get("triggeredBy"), null, callbackMeta);

            if (context.isAuthenticated()) {
                authorizationListener.authorized();
                return USER_AUTHENTICATED_SUCCESSFULLY_PLEASE_PROCEED_WITH_REQUEST;
            }

            return USER_AUTHENTICATED_SUCCESSFULLY + " Please save the changes.";

        } catch (Exception cause) {
            // Note: Do NOT re-throw MustAuthorizeException here. The token was already
            // stored by sdk.handleCallback() above. If hasUserAccess() fails (e.g. AADSTS65001
            // transient consent propagation delay), the token IS persisted and the next
            // Test Connection retry will succeed. Re-throwing MAE would prevent the callback
            // from returning a response and break the SDK auth tab retry flow.
            LOGGER.error("Exception occurred during authorization: {}", cause.getMessage(), cause);
            if (cause instanceof InterruptedException) Thread.currentThread().interrupt();
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "OAuth Callback",
                    cause.getMessage(), cause.getClass().getSimpleName(), callbackMeta);
            return ERROR_OCCURRED_DURING_AUTHORIZATION + " " + cause.getMessage();
        } finally {
            if (authContextId != null) {
                LOGGER.debug("Cleaning up attribute store for authContextId: {}", authContextId);
                officeExcelAttributeStore.remove(authContextId);
            }
        }
    }

    // ── SDK Auth Tab — pre-built HTML (no React/Node build) ───────────────────

    /**
     * Serves the pre-built auth tab HTML from the microsoft-auth-sdk JAR.
     * Replaces the old React-built UI served under /docs/.
     */
    @GET
    @Path("/authentication/{subPath:.*}")
    public InputStream customTabs(@PathParam("subPath") String subPath) {
        return getClass().getClassLoader().getResourceAsStream("auth/index.html");
    }

    // ── SDK Extension Metadata endpoint (consumed by auth tab UI) ─────────────

    @GET
    @Path("/msauth/metadata/{extensionId}")
    @Produces(javax.ws.rs.core.MediaType.APPLICATION_JSON)
    public String getExtensionMetadata(@PathParam("extensionId") String extensionId) {
        ExtensionRegistry.initialize();
        return ExtensionRegistry.get(extensionId)
                .map(m -> m.toUiJson())
                .orElse("{\"error\":\"Extension not found: " + extensionId + "\"}");
    }

    // ── Credentials ────────────────────────────────────────────────────────────

    @GET
    @javax.ws.rs.Path("/getCredentials")
    @Produces("text/plain")
    public String getCredentials(@QueryParam(AUTH_TYPE) String authType,
                                 @QueryParam(AUTH_FLOW) String authFlow) {
        LOGGER.info("Loading attributes for auth-type: {}, auth-flow: {} for invoker {}", authType, authFlow, invokerId);
        final OfficeExcelAttributes loadedAttributes = officeExcelAttributeStore.load(invokerId);
        if (loadedAttributes == null || authType == null || !authType.equals(loadedAttributes.getAuthType())) {
            return "";
        }
        // Use publicAuthMode ("own" = User Auth, "krista" = Quick Setup) to discriminate panels.
        // Backward-compat: credentials saved before publicAuthMode was introduced have it null;
        // derive the effective mode from the stored authFlow in that case.
        if (PUBLIC.equals(authType)) {
            String storedMode = loadedAttributes.getPublicAuthMode();
            if (storedMode == null) {
                storedMode = AUTH_FLOW_PRIVATE_USER.equals(loadedAttributes.getAuthFlow()) ? "own" : "krista";
            }
            if (authFlow != null) {
                if ("AUTH_CODE_PRIVATE".equals(authFlow) && !"own".equals(storedMode)) {
                    return "";
                }
                if ("AUTH_CODE_PUBLIC".equals(authFlow) && "own".equals(storedMode)) {
                    return "";
                }
            } else if (!"own".equals(storedMode)) {
                // Null authFlow from older UI — only serve if stored mode is "own" (backward compat).
                // Quick Setup (krista) data must never leak without explicit authFlow matching.
                return "";
            }
        }
        Map<String, Object> attributesMap = new HashMap<>(loadedAttributes.toMap());
        // Belt-and-suspenders: strip Krista-managed credentials from the Quick Setup response
        // so they never reach the browser, even if toMap() were to include them.
        if (PUBLIC.equals(loadedAttributes.getAuthType())) {
            String storedMode = loadedAttributes.getPublicAuthMode();
            if (storedMode == null) {
                storedMode = AUTH_FLOW_PRIVATE_USER.equals(loadedAttributes.getAuthFlow()) ? "own" : "krista";
            }
            if (!"own".equals(storedMode)) {
                attributesMap.remove(CLIENT_ID);
                attributesMap.remove(CLIENT_SECRET);
            }
        }
        return Constants.GSON.toJson(attributesMap);
    }

    @GET
    @javax.ws.rs.Path("/getAuthKey")
    @Produces("text/plain")
    public String getAuthKey() {
        LOGGER.info("Loading stored authentication key.");
        final OfficeExcelAttributes attributes = officeExcelAttributeStore.load(invokerId);
        if (attributes != null) {
            final String authType = attributes.getAuthType();
            return Constants.GSON.toJson(Objects.requireNonNullElse(authType, PUBLIC));
        } else {
            return "";
        }
    }

    // ── Test Connection ────────────────────────────────────────────────────────

    @POST
    @Path("/testConnection")
    @Produces("text/plain")
    public String testConnection(JsonObject authPayload) {
        auditService.startRequest(invokerId, "Test Connection");
        LOGGER.info("Testing connection started.");
        OfficeExcelAttributes officeExcelAttributes = OfficeExcelAttributes.create(authPayload, baseRoutingUrl);
        Map<String, String> meta = new HashMap<>();
        if (officeExcelAttributes.getEmail() != null) meta.put("triggeredBy", officeExcelAttributes.getEmail());

        String authContextId = providerFactory.createAttributes(officeExcelAttributes);
        String authUrl = null;

        try {
            LOGGER.info("Validating connection with Graph Service Client.");
            GraphServiceClientProvider testProvider = providerFactory.create(authContextId);
            if (AUTH_FLOW_CLIENT_CREDS.equals(officeExcelAttributes.getAuthFlow())) {
                // S2S app-only token: /me requires a delegated token.
                // /users/{email} requires User.Read.All which is NOT a required S2S permission.
                // /users/{email}/drive requires only Files.ReadWrite.All — the same permission
                // needed for actual catalog requests, so this validates the right permission set.
                testProvider.getGraphServiceClientForAdmin()
                        .users(officeExcelAttributes.getEmail())
                        .drive()
                        .buildRequest().get();
            } else {
                testProvider.getGraphServiceClientForAdmin().me()
                        .buildRequest().select("id,displayName,mail").get();
            }
            LOGGER.info("Test connection successful.");
            auditService.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", "Test Connection",
                    "Connection verified for " + officeExcelAttributes.getEmail(), null, meta);
            return createTestConnectionResponse(true, null, null);
        } catch (GraphServiceException cause) {
            LOGGER.error("Graph API error: {}", cause.getMessage(), cause);
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Test Connection",
                    cause.getMessage(), cause.getClass().getSimpleName(), meta);
            return createTestConnectionResponse(false, "An error occurred during test connection.", null);
        } catch (MustAuthorizeException cause) {
            LOGGER.warn("Authorization required, generating authorization URL.");
            if (AUTH_FLOW_CLIENT_CREDS.equals(officeExcelAttributes.getAuthFlow())) {
                // S2S uses no OAuth redirect — surface the credential validation error directly.
                LOGGER.warn("testConnection: CLIENT_CREDENTIALS flow — credential validation failed: {}", cause.getMessage());
                auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Test Connection",
                        cause.getMessage(), "MustAuthorizeException", meta);
                return createTestConnectionResponse(false, cause.getMessage(), null);
            }
            String state = createStateParameter(cause, officeExcelAttributes);
            ITokenStore tokenStore = createTokenStore();
            String baseAuthUrl = officeExcelAttributes.toSdk(tokenStore).buildAuthorizationUrl(state);
            LOGGER.info("testConnection: SDK-generated auth URL scope check — contains Files.ReadWrite: {}",
                    baseAuthUrl.contains("Files.ReadWrite"));

            // The platform SDK for the Quick Setup (Public) flow may generate the auth URL using a
            // shared/platform-level scope set (e.g. email scopes from other extensions registered
            // under the same client_id) instead of this extension's Files.ReadWrite scope.
            // If Files.ReadWrite is missing, the user will silently consent to wrong scopes, the
            // callback fires immediately (no consent screen shown), the token lacks Files.ReadWrite,
            // and MSAL4J fails with AADSTS65001 on every subsequent call.
            // Fix: override the scope to guarantee Files.ReadWrite is always in the consent request.
            if (!baseAuthUrl.contains("Files.ReadWrite")) {
                LOGGER.warn("testConnection: Files.ReadWrite missing from SDK auth URL — overriding scope with extension-specific scopes");
                baseAuthUrl = baseAuthUrl.replaceFirst("scope=[^&]*",
                        "scope=openid%20offline_access%20Files.ReadWrite%20User.Read");
            }

            // Replace the SDK's default prompt value with prompt=consent so Microsoft always shows
            // the full consent screen. The SDK adds prompt=select_account; we replace it rather than
            // append to avoid AADSTS9000411 (duplicate parameter).
            authUrl = baseAuthUrl.contains("prompt=")
                    ? baseAuthUrl.replaceFirst("prompt=[^&]*", "prompt=consent")
                    : baseAuthUrl + "&prompt=consent";
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Test Connection",
                    "Authentication required — redirecting to sign-in", "MustAuthorizeException", meta);
            return createTestConnectionResponse(false, AUTHORIZATION_PROMPT, authUrl);
        } catch (IOException e) {
            LOGGER.error("IO error during test connection: {}", e.getMessage(), e);
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Test Connection",
                    e.getMessage(), e.getClass().getSimpleName(), meta);
            throw new RuntimeException(e);
        } finally {
            if (authUrl == null) {
                LOGGER.info("Cleaning up authentication context.");
                officeExcelAttributeStore.remove(authContextId);
            }
        }
    }

    // ── Save Credentials ───────────────────────────────────────────────────────

    @POST
    @Path("/saveCredentials")
    @Produces("text/plain")
    public String saveCredentials(JsonObject authPayload) {
        auditService.startRequest(invokerId, "Save Credentials");
        OfficeExcelAttributes attributes = OfficeExcelAttributes.create(authPayload, baseRoutingUrl);
        Map<String, String> meta = new HashMap<>();
        if (attributes.getEmail() != null) meta.put("triggeredBy", attributes.getEmail());

        String authContextId = officeExcelAttributeStore.save(attributes);
        try {
            GraphServiceClientProvider saveProvider = providerFactory.create(attributes);
            if (AUTH_FLOW_CLIENT_CREDS.equals(attributes.getAuthFlow())) {
                // S2S: /users/{email}/drive validates Files.ReadWrite.All — the required S2S permission
                saveProvider.getGraphServiceClientForAdmin()
                        .users(attributes.getEmail())
                        .drive()
                        .buildRequest().get();
            } else {
                saveProvider.getGraphServiceClientForAdmin().me()
                        .buildRequest().select("id,displayName,mail").get();
            }
            // Clean up old token when switching email or auth mode — prevents orphaned
            // refresh tokens that can never be reached by the new configuration.
            OfficeExcelAttributes oldAttributes = officeExcelAttributeStore.load(invokerId);
            if (oldAttributes != null && oldAttributes.getEmail() != null) {
                String oldTokenKey = MicrosoftAuthSdk.buildTokenKey(
                        oldAttributes.getEmail(), oldAttributes.getClientId(), oldAttributes.getAuthType());
                String newTokenKey = MicrosoftAuthSdk.buildTokenKey(
                        attributes.getEmail(), attributes.getClientId(), attributes.getAuthType());
                if (!oldTokenKey.equals(newTokenKey)) {
                    refreshTokenStore.remove(oldTokenKey);
                    LOGGER.info("Removed old token for key: {} (switching to: {})", oldTokenKey, newTokenKey);
                }
            }
            boolean isSaved = officeExcelAttributeStore.save(attributes, invokerId);
            if (isSaved) {
                auditService.resetAuthCounters();
                auditService.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", "Save Credentials",
                        "Configuration saved successfully", null, meta);
                return Constants.GSON.toJson(new AuthenticationResponse(true, null, null));
            } else {
                auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Save Credentials",
                        FAILED_TO_SAVE_ATTRIBUTES, null, meta);
                return Constants.GSON.toJson(new AuthenticationResponse(false, FAILED_TO_SAVE_ATTRIBUTES, null));
            }
        } catch (MustAuthorizeException e) {
            if (AUTH_FLOW_CLIENT_CREDS.equals(attributes.getAuthFlow())) {
                // S2S has no OAuth redirect — surface the credential error directly
                LOGGER.error("S2S credentials validation failed during save: {}", e.getMessage());
                auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Save Credentials",
                        e.getMessage(), "MustAuthorizeException", meta);
                return Constants.GSON.toJson(new AuthenticationResponse(false, e.getMessage(), null));
            }
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Save Credentials",
                    "Authentication required — redirecting to sign-in", "MustAuthorizeException", meta);
            throw e;
        } catch (Exception cause) {
            LOGGER.error(FAILED_TO_SAVE_ATTRIBUTES);
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Save Credentials",
                    cause.getMessage(), cause.getClass().getSimpleName(), meta);
            return Constants.GSON.toJson(new AuthenticationResponse(false, FAILED_TO_SAVE_ATTRIBUTES, null));
        } finally {
            officeExcelAttributeStore.remove(authContextId);
        }
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    private GraphServiceClientProvider getCallbackProvider(String authContextId, String key) {
        if (key != null && key.startsWith(WS_CONTACT)) {
            OfficeExcelAttributes attrs = new OfficeExcelAttributes(null, null, null, null, PUBLIC, baseRoutingUrl);
            return providerFactory.create(attrs);
        }
        if (authContextId == null) {
            return providerFactory.create();
        }
        OfficeExcelAttributes effectiveAttributes = officeExcelAttributeStore.load(authContextId);
        return providerFactory.create(effectiveAttributes);
    }

    private String extractEmailFromAccessToken(String token) {
        try {
            String[] jwtParts = token.split("\\.");
            if (jwtParts.length >= 2) {
                int mod = jwtParts[1].length() % 4;
                String padded = mod > 0 ? jwtParts[1] + "=".repeat(4 - mod) : jwtParts[1];
                byte[] decoded = Base64.getUrlDecoder().decode(padded);
                Map<?, ?> claims = Constants.GSON.fromJson(new String(decoded), Map.class);
                if (claims.containsKey("preferred_username")) return (String) claims.get("preferred_username");
                if (claims.containsKey("upn")) return (String) claims.get("upn");
                if (claims.containsKey("email")) return (String) claims.get("email");
            }
        } catch (Exception e) {
            LOGGER.warn("Failed to extract email from access token: {}", e.getMessage());
        }
        return null;
    }

    private String createTestConnectionResponse(boolean isSuccess, String errorMessage, String url) {
        if (isSuccess) {
            return Constants.GSON.toJson(new AuthenticationResponse(true, null, null));
        } else if (url != null) {
            return Constants.GSON.toJson(new AuthenticationResponse(false, errorMessage, url));
        } else {
            return Constants.GSON.toJson(new AuthenticationResponse(false,
                    Objects.requireNonNullElse(errorMessage, "Unknown Error"), null));
        }
    }

    private String createStateParameter(MustAuthorizeException cause, OfficeExcelAttributes officeExcelAttributes) {
        String userId;
        String authContextId = null;

        if (cause != null) {
            userId = (String) cause.getDetails().get(0).getValue();
            Optional<NamedValuedField> authContextIdField = cause.getDetails().stream()
                    .filter(f -> Objects.equals(f.getName(), AUTH_CONTEXT_ID))
                    .findFirst();
            if (authContextIdField.isPresent()) {
                authContextId = (String) authContextIdField.get().getValue();
            }
        } else {
            userId = officeExcelAttributes.getEmailId() != null ? officeExcelAttributes.getEmailId() : "test-user";
        }

        String state = userId;
        if (authContextId != null) {
            state += HASH + authContextId;
        }
        String forwardPath = officeExcelAttributes.getForwardPath();
        if (!forwardPath.isEmpty()) {
            state += HASH + forwardPath;
        } else if (AUTH_FLOW_PRIVATE_USER.equals(officeExcelAttributes.getAuthFlow())) {
            // For User Authentication (own Azure AD app), include the callback endpoint in the
            // state so the SDK can find it during handleCallback() to complete the token exchange.
            state += HASH + officeExcelAttributes.getCallbackEndPoint();
        }
        return state;
    }

    private String getAuthContextId(String state) {
        if (state == null || state.isEmpty()) {
            return null;
        }
        // Auth context IDs start with AUTH_CONTEXT_ID prefix
        int authContextIdIndex = state.indexOf(AUTH_CONTEXT_ID);
        if (authContextIdIndex != -1) {
            String authContextIdToEnd = state.substring(authContextIdIndex);
            int hashIndex = authContextIdToEnd.indexOf(HASH);
            return hashIndex != -1
                    ? authContextIdToEnd.substring(0, hashIndex)
                    : authContextIdToEnd;
        }
        // Fallback: positional extraction
        String[] parts = state.split(HASH);
        if (parts.length >= 2) {
            String potentialAuthContextId = parts[1];
            if (!potentialAuthContextId.startsWith("/") && !potentialAuthContextId.isBlank()) {
                return potentialAuthContextId;
            }
        }
        return null;
    }

    private ITokenStore createTokenStore() {
        return new ITokenStore() {
            @Override public void save(String k, String v) {
                try { refreshTokenStore.put(k, v); } catch (IOException e) { throw new RuntimeException(e); }
            }
            @Override public String get(String k) {
                try { return refreshTokenStore.get(k); } catch (IOException e) { throw new RuntimeException(e); }
            }
            @Override public void delete(String k) { refreshTokenStore.remove(k); }
        };
    }
}
