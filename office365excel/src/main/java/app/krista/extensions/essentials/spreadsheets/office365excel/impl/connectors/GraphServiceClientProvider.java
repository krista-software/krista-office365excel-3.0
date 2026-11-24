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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.AuthErrorRule;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.field.NamedValuedField;
import com.google.gson.JsonObject;
import com.microsoft.aad.msal4j.*;
import com.microsoft.graph.authentication.IAuthenticationProvider;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.httpcore.HttpClients;
import com.microsoft.graph.models.User;
import com.microsoft.graph.requests.DriveRequestBuilder;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.requests.WorkbookWorksheetRequestBuilder;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import org.jetbrains.annotations.NotNull;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.MicrosoftAuthErrorConstants.AUTH_ERROR_RULES;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;

/**
 * Provider class for Microsoft Graph API client instances.
 * This class handles authentication, token management, and client creation for Office365 Excel operations.
 */
@Service
public class GraphServiceClientProvider {
    public static final String AUTHORIZATION_PROMPT = "Authorize yourself and click Validate Attributes before saving changes to proceed further";
    private static final Logger LOGGER = LoggerFactory.getLogger(GraphServiceClientProvider.class);

    private final RefreshTokenStore refreshTokenStore;
    private final OfficeExcelAttributes officeExcelAttributes;
    private final AuthorizationContext authorizationContext;
    private final OfficeExcelAttributeStore officeExcelAttributestore;
    private final Invoker invoker;
    private final String authContextId;
    private RequestContext requestContext;

    @Inject
    public GraphServiceClientProvider(RefreshTokenStore refreshTokenStore, OfficeExcelAttributes officeExcelAttributes, AuthorizationContext authorizationContext, OfficeExcelAttributeStore officeExcelAttributestore, Invoker invoker, String authContextId) {
        this.refreshTokenStore = refreshTokenStore;
        this.officeExcelAttributes = officeExcelAttributes;
        this.authorizationContext = authorizationContext;
        this.officeExcelAttributestore = officeExcelAttributestore;
        this.invoker = invoker;
        this.authContextId = authContextId;
    }

    public GraphServiceClientProvider(RefreshTokenStore refreshTokenStore, OfficeExcelAttributes officeExcelAttributes, OfficeExcelAttributeStore officeExcelAttributestore, Invoker invoker,
                                      RequestContext requestContext, AuthorizationContext authorizationContext,
                                      String authContextId) {
        this.refreshTokenStore = refreshTokenStore;
        this.officeExcelAttributes = officeExcelAttributes;
        this.officeExcelAttributestore = officeExcelAttributestore;
        this.invoker = invoker;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.authContextId = authContextId;
    }

    /**
     * Updates the request context for this provider.
     *
     * @param requestContext The new request context to use
     */
    public void setRequestContext(RequestContext requestContext) {
        this.requestContext = requestContext;
    }

    /**
     * Gets the Office 365 Excel attributes used by this provider.
     *
     * @return The Office 365 Excel attributes
     */
    public OfficeExcelAttributes getOfficeExcelAttributes() {
        return officeExcelAttributes;
    }

    /**
     * Gets a REST client for the specified drive.
     * Uses the current user context to determine authentication method.
     *
     * @param driveId The ID of the drive to access
     * @return A REST client configured for the specified drive
     */
    public RestClient getRestClient(String driveId) throws IOException {
        checkUserToken(null, null);
        return getRestClient(driveId, !requestContext.invokeAsUser());
    }

    public GraphServiceClient<Request> getGraphServiceClientForAdmin() throws IOException {
        OfficeExcelAttributes effectiveAttrs = officeExcelAttributes != null
                ? officeExcelAttributes
                : officeExcelAttributestore.load(invoker.getInvokerId());
        if (effectiveAttrs != null && AUTH_FLOW_CLIENT_CREDS.equals(effectiveAttrs.getAuthFlow())) {
            return getGraphServiceClientForClientCredentials(effectiveAttrs);
        }
        return getGraphServiceClient(true, null);
    }

    /**
     * Validates the refresh token for the specified account.
     * Throws a MustAuthorizeException if the token is invalid or missing.
     *
     * @param accountId The account identifier.
     * @throws MustAuthorizeException If refresh token is null then it will throw MustAuthorize Exception.
     */
    public void validateToken(String accountId) {
        try {
            getGraphServiceClient(false, accountId);
        } catch (IOException e) {
            LOGGER.warn("validateToken: error reading token store for account={} — {}", accountId, e.getMessage());
        }
    }

    /**
     * Gets a GraphServiceClient based on the provided authentication parameters.
     * Handles token retrieval and error cases.
     *
     * @param useEmail  Whether to use the admin email for authentication
     * @param accountID The account ID to use if useEmail is false
     * @return A GraphServiceClient configured for the specified authentication
     * @throws RuntimeException If authentication fails
     */
    private GraphServiceClient<Request> getGraphServiceClient(boolean useEmail, String accountID) throws IOException {
        // S2S (Client Credentials): no refresh token — acquire app-only token directly from Azure AD
        OfficeExcelAttributes effectiveAttrs = officeExcelAttributes != null
                ? officeExcelAttributes
                : officeExcelAttributestore.load(invoker.getInvokerId());
        if (effectiveAttrs != null && AUTH_FLOW_CLIENT_CREDS.equals(effectiveAttrs.getAuthFlow())) {
            return getGraphServiceClientForClientCredentials(effectiveAttrs);
        }
        try {
            String userId = getUserId(useEmail, accountID);
            String refreshTokenStoreKey = getRefTokenStoreKey(userId);
            String refreshToken = refreshTokenStore.get(refreshTokenStoreKey);
            if (refreshToken == null) {
                throw createMustAuthorizationException(refreshTokenStoreKey, false);
            }
            return getGraphServiceClient(refreshTokenStoreKey, refreshToken);
        } catch (RuntimeException | IOException cause) {
            throw cause;
        }
    }

    /**
     * Creates a new GraphServiceClient instance using the provided user ID and refresh token.
     * This method is used internally to obtain a GraphServiceClient object for authorized access to Office 365 Excel
     * resources on behalf of the specified user.
     *
     * @param userId       The unique identifier of the user for whom the GraphServiceClient should be created.
     * @param refreshToken The refresh token associated with the user's account.
     * @return A GraphServiceClient object authorized to access Office 365 Excel resources on behalf of the specified user.
     * @throws MustAuthorizeException If authorization fails due to invalid refresh token or other issues.
     */
    private GraphServiceClient<Request> getGraphServiceClient(String userId, String refreshToken) {
        try {
            OfficeExcelAttributes load = officeExcelAttributestore.load(invoker.getInvokerId());
            String[] scopes = Constants.SCOPE.split(Constants.SCOPE_SEPARATOR);
            Set<String> scopeSet = Arrays.stream(scopes).collect(Collectors.toSet());
            RefreshTokenParameters refreshTokenParameters = RefreshTokenParameters.builder(scopeSet, refreshToken).build();

            ConfidentialClientApplication confidentialClientApplication = buildConfidentialClientApp(load, officeExcelAttributes);

            IAuthenticationResult authenticationResult = confidentialClientApplication.acquireToken(refreshTokenParameters).get();
            final String cachedTokenContent = confidentialClientApplication.tokenCache().serialize();
            updateRefreshToken(userId, Constants.GSON.fromJson(cachedTokenContent, JsonObject.class));

            IAuthenticationProvider authProvider = new GraphServiceClientAuthenticationProvider(authenticationResult.accessToken());
            // createSession on large workbooks can take 25-30 s; the default OkHttp
            // read timeout (10 s) fires before the response arrives, causing the
            // session manager to return null and the subsequent sessionless call to
            // get a 403 "Could not obtain a WAC access token".
            OkHttpClient httpClient = HttpClients.createDefault(authProvider)
                    .newBuilder()
                    .readTimeout(60, TimeUnit.SECONDS)
                    .callTimeout(120, TimeUnit.SECONDS)
                    .build();
            return GraphServiceClient.builder()
                    .httpClient(httpClient)
                    .buildClient();
        } catch (ExecutionException cause) {
            // Unwrap the MSAL4J exception to get the actual Azure AD error code (e.g. AADSTS65001).
            // ExecutionException wraps the real cause — getCause() holds the MsalServiceException.
            Throwable msalCause = cause.getCause() != null ? cause.getCause() : cause;
            String errorMsg = msalCause.getMessage();
            LOGGER.error("getGraphServiceClient: MSAL token acquisition failed for userId={}: {}", userId, errorMsg);
            if (errorMsg != null) {
                for (AuthErrorRule rule : AUTH_ERROR_RULES) {
                    if (rule.matches(errorMsg)) {
                        if (rule.shouldRemoveToken()) {
                            refreshTokenStore.remove(userId);
                        }
                        throw new MustAuthorizeException(rule.getUserMessage(), createAuthDetails(userId));
                    }
                }
            }
            throw createMustAuthorizationException(userId, true);
        } catch (ClientException | MalformedURLException cause) {
            throw createMustAuthorizationException(userId, true);
        } catch (InterruptedException cause) {
            Thread.currentThread().interrupt();
            throw createMustAuthorizationException(userId, true);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public GraphServiceClient getGraphServiceClient(RequestContext threadLocalRequestContext) throws IOException {
        return getGraphServiceClient(!threadLocalRequestContext.invokeAsUser(), null);
    }

    private ConfidentialClientApplication buildConfidentialClientApp(OfficeExcelAttributes load, OfficeExcelAttributes officeExcelAttributes)
            throws MalformedURLException {
        OfficeExcelAttributes effectiveAttributes = (officeExcelAttributes == null) ? load : officeExcelAttributes;

        IClientCredential clientCredential = ClientCredentialFactory.createFromSecret(effectiveAttributes.getClientSecret());

        String authority = determineAuthorityUrl(effectiveAttributes.getAuthType(), effectiveAttributes.getTenantId());

        return ConfidentialClientApplication.builder(effectiveAttributes.getClientId(), clientCredential)
                .authority(authority)
                .build();
    }

    /**
     * Acquires an app-only access token using the OAuth 2.0 Client Credentials grant.
     * Used exclusively for Server-to-Server authentication (no user sign-in required).
     */
    private IAuthenticationResult acquireClientCredentialsToken(OfficeExcelAttributes attrs)
            throws MalformedURLException, ExecutionException, InterruptedException {
        Set<String> scopes = Collections.singleton(Constants.CLIENT_CREDENTIALS_SCOPE);
        ClientCredentialParameters parameters = ClientCredentialParameters.builder(scopes).build();
        ConfidentialClientApplication app = buildConfidentialClientApp(null, attrs);
        return app.acquireToken(parameters).get();
    }

    /**
     * Builds a GraphServiceClient using the Client Credentials flow (Server-to-Server).
     * Does not require a refresh token — obtains a fresh app-only token from Azure AD.
     */
    private GraphServiceClient<Request> getGraphServiceClientForClientCredentials(OfficeExcelAttributes attrs) {
        try {
            IAuthenticationResult result = acquireClientCredentialsToken(attrs);
            IAuthenticationProvider authProvider = new GraphServiceClientAuthenticationProvider(result.accessToken());
            OkHttpClient httpClient = HttpClients.createDefault(authProvider)
                    .newBuilder()
                    .readTimeout(60, TimeUnit.SECONDS)
                    .callTimeout(120, TimeUnit.SECONDS)
                    .build();
            return GraphServiceClient.builder().httpClient(httpClient).buildClient();
        } catch (ExecutionException e) {
            LOGGER.error("Client credentials token acquisition failed: {}", e.getMessage(), e);
            Throwable cause = e.getCause() != null ? e.getCause() : e;
            throw new MustAuthorizeException(
                    "Invalid credentials. Please verify your Client ID, Client Secret, and Tenant ID. ("
                            + cause.getMessage() + ")", Collections.emptyList());
        } catch (MalformedURLException e) {
            LOGGER.error("Malformed authority URL for tenant: {}", e.getMessage(), e);
            throw new MustAuthorizeException("Invalid Tenant ID format.", Collections.emptyList());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new MustAuthorizeException("Authentication request interrupted.", Collections.emptyList());
        }
    }

    /**
     * Builds a RestClient using the Client Credentials flow (Server-to-Server).
     * Does not require a refresh token — obtains a fresh app-only token from Azure AD.
     */
    private RestClient getRestClientForClientCredentials(String driveId, OfficeExcelAttributes attrs) {
        try {
            IAuthenticationResult result = acquireClientCredentialsToken(attrs);
            return new RestClient(result.accessToken(), driveId);
        } catch (ExecutionException e) {
            LOGGER.error("Client credentials REST token acquisition failed: {}", e.getMessage(), e);
            Throwable cause = e.getCause() != null ? e.getCause() : e;
            throw new MustAuthorizeException(
                    "Invalid credentials for server-to-server REST access. ("
                            + cause.getMessage() + ")", Collections.emptyList());
        } catch (MalformedURLException e) {
            LOGGER.error("Malformed authority URL: {}", e.getMessage(), e);
            throw new MustAuthorizeException("Invalid Tenant ID format.", Collections.emptyList());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new MustAuthorizeException("REST authentication request interrupted.", Collections.emptyList());
        }
    }

    private String determineAuthorityUrl(String authType, String tenantId) {
        if (PRIVATE.equals(authType)) {
            return Constants.AUTHORITY + tenantId;
        } else if (PUBLIC.equals(authType)) {
            if (tenantId != null && !tenantId.isBlank()) {
                // User Authentication with own Azure AD app — tenant-specific authority required.
                // Using the generic organizations/ endpoint with a single-tenant app's refresh token
                // causes AADSTS65001 (consent not found in the "organizations" virtual tenant).
                return Constants.AUTHORITY + tenantId;
            }
            // Quick Setup — Krista shared multi-tenant app, organizations/ endpoint is correct.
            return Constants.ORG_AUTHORITY;
        } else {
            throw new IllegalArgumentException("No authentication type found.");
        }
    }

    /**
     * Gets a REST client for the specified drive with explicit authentication parameters.
     * Handles token retrieval and error cases.
     *
     * @param driveId  The ID of the drive to access
     * @param useEmail Whether to use the admin email for authentication
     * @return A REST client configured for the specified drive and authentication
     */
    private RestClient getRestClient(String driveId, boolean useEmail) throws IOException {
        // S2S (Client Credentials): no refresh token — acquire app-only token directly from Azure AD
        OfficeExcelAttributes effectiveAttrs = officeExcelAttributes != null
                ? officeExcelAttributes
                : officeExcelAttributestore.load(invoker.getInvokerId());
        if (effectiveAttrs != null && AUTH_FLOW_CLIENT_CREDS.equals(effectiveAttrs.getAuthFlow())) {
            return getRestClientForClientCredentials(driveId, effectiveAttrs);
        }
        try {
            String userId = getUserId(useEmail, null);
            String refreshTokenStoreKey = getRefTokenStoreKey(userId);
            String refreshToken = refreshTokenStore.get(refreshTokenStoreKey);
            if (refreshToken == null) {
                throw createMustAuthorizationException(refreshTokenStoreKey, false);
            }
            return getRestClient(driveId, refreshTokenStoreKey, refreshToken);
        } catch (RuntimeException | IOException cause) {
            LOGGER.error("Failed to get graph client {}", cause.getClass().getName());
            throw cause;
        }
    }

    /**
     * Creates a REST client for the specified drive using the provided authentication parameters.
     * Handles token acquisition and error cases.
     *
     * @param driveId      The ID of the drive to access
     * @param userId       The user ID for authentication
     * @param refreshToken The refresh token for authentication
     * @return A REST client configured for the specified drive and authentication
     */
    private RestClient getRestClient(String driveId, String userId, String refreshToken) {
        try {
            OfficeExcelAttributes load = officeExcelAttributestore.load(invoker.getInvokerId());
            String[] scopes = Constants.SCOPE.split(Constants.SCOPE_SEPARATOR);
            Set<String> scopeSet = Arrays.stream(scopes).collect(Collectors.toSet());
            RefreshTokenParameters refreshTokenParameters = RefreshTokenParameters.builder(scopeSet, refreshToken).build();
            ConfidentialClientApplication confidentialClientApplication = buildConfidentialClientApp(load, officeExcelAttributes);
            IAuthenticationResult authenticationResult = confidentialClientApplication.acquireToken(refreshTokenParameters).get();
            return new RestClient(authenticationResult.accessToken(), driveId);
        } catch (IOException | ExecutionException | InterruptedException cause) {
            Thread.currentThread().interrupt();
            throw createMustAuthorizationException(userId, true);
        }
    }

    /**
     * Updates the refresh token for a user based on the token cache data.
     * Extracts the new refresh token from the JSON object and stores it in the token store.
     *
     * @param userId     The user ID to update the token for
     * @param jsonObject The JSON object containing the token cache data
     */
    private void updateRefreshToken(String userId, JsonObject jsonObject) throws IOException {
        JsonObject refreshToken = jsonObject.get("RefreshToken").getAsJsonObject();
        String firstKey = refreshToken.keySet().iterator().next();
        JsonObject tokenData = refreshToken.getAsJsonObject(firstKey);
        String secret = tokenData.get("secret").getAsString();
        LOGGER.info("Refresh token logs :::::: {}", secret);
        refreshTokenStore.put(userId, secret);
    }

    /**
     * Builds a GraphServiceClient object for performing operations on behalf of the current user or a specific user.
     * <p>
     * This method provides multiple options for obtaining a GraphServiceClient:
     * - If `useEmail` is true and `accountID` is provided, the method retrieves the user by email and creates a client for them.
     * - If `useEmail` is true and `accountID` is not provided, the method uses the currently logged-in user's email to create a client.
     * - If `useEmail` is null, the method checks `RequestContext.invokeAsUser`:
     * - If `invokeAsUser` is true, it creates a client for the currently logged-in user.
     * - If `invokeAsUser` is false, it throws an IllegalStateException because a user context is required.
     *
     * @param useEmail  Whether to use the provided email (if true) or the currently logged-in user (if null).
     * @param accountID The email address of the specific user (relevant only if `useEmail` is true).
     * @return A GraphServiceClient object authorized to access Office 365 Excel resources on behalf of the specified user.
     * @throws IOException           If any I/O error occurs while retrieving user information.
     * @throws IllegalStateException If a user context is required but not available.
     */
    public GraphServiceClient<Request> getUserRequestBuilder(Boolean useEmail, String accountID) throws IOException {
        checkUserToken(useEmail, accountID);
        if (useEmail != null) {
            return getGraphServiceClient(useEmail, useEmail ? null : accountID);
        }
        return getGraphServiceClient(!requestContext.invokeAsUser(), null);
    }


    /**
     * Creates a MustAuthorizeException with appropriate details for authentication failures.
     * Configures the exception with user ID and context information.
     *
     * @param userId           The user ID that failed authentication
     * @param reAuthentication Whether this is a re-authentication attempt
     * @return A configured MustAuthorizeException
     */
    private MustAuthorizeException createMustAuthorizationException(String userId, boolean reAuthentication) {
        LOGGER.info(reAuthentication ? Constants.GOT_ERROR_FOR_AUTHENTICATION_SO_SENDING_FOR_RE_AUTHENTICATION :
                Constants.GOT_ERROR_FOR_AUTHENTICATION_SO_SENDING_FOR_AUTHENTICATION);
        List<NamedValuedField> details = new ArrayList<>();
        NamedValuedField userIdField = new NamedValuedField(Constants.USER_ID, Constants.TEXT, userId, new HashMap<>(), new HashMap<>());
        details.add(userIdField);
        if (authContextId != null) {
            NamedValuedField contextIdField = new NamedValuedField(Constants.AUTH_CONTEXT_ID, Constants.TEXT,
                    authContextId, new HashMap<>(), new HashMap<>());
            details.add(contextIdField);
        }
        return new MustAuthorizeException(reAuthentication ?
                Constants.REFRESH_TOKEN_IS_NO_LONGER_VALID_PLEASE_REAUTHORIZE_YOURSELF :
                AUTHORIZATION_PROMPT, details);

    }

    /**
     * Gets the user ID based on the authentication context.
     * Determines whether to use the admin email or the current user's account ID.
     *
     * @param calledFromValidateAttributes Whether this is called from validateAttributes
     * @param accountID                    The account ID to use if not called from validateAttributes
     * @return The user ID to use for authentication
     */
    private String getUserId(boolean calledFromValidateAttributes, String accountID) {
        String userId;
        if (calledFromValidateAttributes) {
            // Prefer the email from the current operation's attributes (e.g. Public test-connection
            // payload). Only fall back to the permanently stored invokerId attributes when the
            // current attributes carry no email — prevents Quick Setup (Private) admin email from
            // leaking into the User Authentication (Public) OAuth flow.
            String emailFromCurrentAttrs = (officeExcelAttributes != null) ? officeExcelAttributes.getEmail() : null;
            if (emailFromCurrentAttrs != null) {
                userId = emailFromCurrentAttrs;
            } else {
                OfficeExcelAttributes load = officeExcelAttributestore.load(invoker.getInvokerId());
                userId = (load != null) ? load.getEmail() : null;
            }
        } else {
            userId = (accountID != null) ? accountID : authorizationContext.getAuthorizedAccount().getAccountId();
        }
        if (userId == null) {
            throw new IllegalStateException(Constants.FAILED_TO_GET_ACCOUNT);
        }
        return userId;
    }

    @NotNull
    private String getRefTokenStoreKey(String userId) {
        OfficeExcelAttributes load = officeExcelAttributestore.load(invoker.getInvokerId());
        if (userId.startsWith(Constants.WS_CONTACT)) {
            return userId;
        } else {
            return userId + Constants.UNDER_SCORE + (officeExcelAttributes == null ? load.getClientId() : officeExcelAttributes.getClientId()) + Constants.UNDER_SCORE + (officeExcelAttributes == null ? load.getAuthType() : officeExcelAttributes.getAuthType());
        }
    }

    /**
     * Authentication provider implementation for Microsoft Graph API.
     * Provides access tokens for Graph API requests.
     */
    public static class GraphServiceClientAuthenticationProvider implements IAuthenticationProvider {

        private final String accessToken;

        /**
         * Constructs an authentication provider with the specified access token.
         *
         * @param accessToken The access token to use for authentication
         */
        public GraphServiceClientAuthenticationProvider(String accessToken) {
            if (accessToken == null || accessToken.isEmpty() || accessToken.isBlank()) {
                throw new IllegalArgumentException("Empty access token provided " + accessToken);
            }
            this.accessToken = accessToken;
        }

        /**
         * Gets the authorization token for a request.
         * Returns the stored access token as a completed future.
         *
         * @param requestUrl The URL of the request
         * @return A CompletableFuture containing the access token
         */
        @NotNull
        @Override
        public CompletableFuture<String> getAuthorizationTokenAsync(@NotNull URL requestUrl) {
            return CompletableFuture.completedFuture(accessToken);
        }
    }

    /**
     * Retrieves the email address of the authenticated admin user.
     * Uses the admin Graph Service Client to query the /me endpoint.
     *
     * @return the email address of the authenticated user, empty string if not found
     * @throws IOException if authentication fails or the API call fails
     */
    public String getAuthenticatedUserEmail() throws IOException {
        User user = getGraphServiceClientForAdmin().me().buildRequest().get();
        return user != null ? user.mail : "";
    }

    public DriveRequestBuilder getDriveRequestBuilder(String driveId, String accountId, Boolean invokeAsUser) throws IOException {
        return StringUtil.isNullOrEmpty(driveId) ? getGraphServiceClient().me().drive() :
                getGraphServiceClientForAsync(invokeAsUser, accountId).me().drives(driveId);
    }

    public DriveRequestBuilder getDriveRequestBuilder(String driveId) {
        try {
            return StringUtil.isNullOrEmpty(driveId) ? getGraphServiceClient().me().drive() : getGraphServiceClient().me().drives(driveId);
        } catch (IOException cause) {
            throw new RuntimeException(cause);
        }
    }

    public DriveRequestBuilder getDriveRequestBuilder(String driveId, RequestContext requestContext) {
        try {
            return StringUtil.isNullOrEmpty(driveId) ? getGraphServiceClient(requestContext).me().drive() : getGraphServiceClient(requestContext).me().drives(driveId);
        } catch (IOException cause) {
            throw new RuntimeException(cause);
        }
    }

    public WorkbookWorksheetRequestBuilder worksheetRequestBuilders(String driveId, String worksheetName, String workbookId) {
        return getDriveRequestBuilder(driveId)
                .items(workbookId)
                .workbook()
                .worksheets(worksheetName);
    }

    /**
     * Handles authentication errors and throws appropriate MustAuthorizeException
     *
     * @param cause                The exception that occurred during authentication
     * @param refreshTokenStoreKey The key used to store the refresh token
     * @throws MustAuthorizeException with appropriate error message based on the cause
     */
    private void handleAuthenticationError(Exception cause, String refreshTokenStoreKey) {
        String errorMessage = cause.getMessage();
        LOGGER.error(" handleAuthenticationError() -> errorMessage ::: {} ", errorMessage);
        if (errorMessage != null) {
            for (AuthErrorRule rule : AUTH_ERROR_RULES) {
                if (rule.matches(errorMessage)) {
                    if (rule.shouldRemoveToken()) {
                        refreshTokenStore.remove(refreshTokenStoreKey);
                    }
                    throw new MustAuthorizeException(rule.getUserMessage(), createAuthDetails(refreshTokenStoreKey));
                }
            }
        }
        throw createMustAuthorizationException(refreshTokenStoreKey, !requestContext.invokeAsUser());
    }

    private List<NamedValuedField> createAuthDetails(String userId) {
        List<NamedValuedField> details = new ArrayList<>();
        NamedValuedField userIdField = new NamedValuedField(USER_ID, TEXT, userId, new HashMap<>(), new HashMap<>());
        details.add(userIdField);
        if (authContextId != null) {
            NamedValuedField contextIdField = new NamedValuedField(AUTH_CONTEXT_ID, TEXT, authContextId, new HashMap<>(), new HashMap<>());
            details.add(contextIdField);
        }
        return details;
    }

    private void checkUserToken(Boolean useEmail, String accountID) {
        String wsContactId = resolveWsContactId(useEmail, accountID);
        if (wsContactId == null) return;
        LOGGER.info("checkUserToken: checking token for wsContactId={}", wsContactId);
        try {
            String token = refreshTokenStore.get(wsContactId);
            if (token == null) {
                LOGGER.info("checkUserToken: no token for wsContactId={} — throwing MAE", wsContactId);
                List<NamedValuedField> details = new ArrayList<>();
                details.add(new NamedValuedField(Constants.CLIENT_USER_ID, Constants.TEXT, wsContactId, new HashMap<>(), new HashMap<>()));
                String invokerId = requestContext != null ? requestContext.getInvokerId() : null;
                if (invokerId != null) {
                    details.add(new NamedValuedField(Constants.USER_ID, Constants.TEXT, invokerId, new HashMap<>(), new HashMap<>()));
                }
                throw new MustAuthorizeException("No credentials configured. Please authenticate via the Authentication tab.", details);
            }
            LOGGER.info("checkUserToken: token found for wsContactId={} — proceeding", wsContactId);
        } catch (MustAuthorizeException e) {
            throw e;
        } catch (IOException e) {
            LOGGER.warn("checkUserToken: error reading token store — {}", e.getMessage());
        }
    }

    private String resolveWsContactId(Boolean useEmail, String accountID) {
        if (accountID != null && accountID.startsWith(WS_CONTACT)) return accountID;
        if (!resolveUseEmail(useEmail)) {
            try {
                String authId = authorizationContext.getAuthorizedAccount().getAccountId();
                if (authId != null && authId.startsWith(WS_CONTACT)) return authId;
            } catch (Exception e) {
                LOGGER.warn("resolveWsContactId: could not resolve from authorizationContext ({})", e.getMessage());
            }
        }
        return null;
    }

    boolean resolveUseEmail(Boolean useEmail) {
        return useEmail != null ? useEmail : !isInvokeAsUser();
    }

    private boolean isInvokeAsUser() {
        return requestContext != null && requestContext.invokeAsUser();
    }

    public GraphServiceClient getGraphServiceClientForAsync(boolean useEmail, String accountId) throws IOException {
        return getGraphServiceClient(useEmail, accountId);
    }

    public GraphServiceClient getGraphServiceClient() throws IOException {
        return getGraphServiceClient(!requestContext.invokeAsUser(), null);
    }

}
