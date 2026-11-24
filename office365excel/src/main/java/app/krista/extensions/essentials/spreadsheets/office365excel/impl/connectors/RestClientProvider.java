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
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.MessageConstants;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.field.NamedValuedField;
import com.microsoft.aad.msal4j.*;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

@Service
public final class RestClientProvider {

    private final OfficeExcelAttributes officeExcelAttributes;
    private final OfficeExcelAttributeStore officeExcelAttributeStore;
    private final Invoker invoker;
    private final RequestContext requestContext;
    private final RefreshTokenStore refreshTokenStore;
    private final AuthorizationContext authorizationContext;
    private static final Logger LOGGER = LoggerFactory.getLogger(RestClientProvider.class);

    /**
     * HK2-managed constructor. {@code OfficeExcelAttributes} is not an HK2 service and
     * cannot be injected — attributes are always resolved at call-time from the attribute store.
     */
    @Inject
    public RestClientProvider(OfficeExcelAttributeStore officeExcelAttributeStore,
                              Invoker invoker, RefreshTokenStore refreshTokenStore,
                              RequestContext requestContext, AuthorizationContext authorizationContext) {
        this.officeExcelAttributes = null;
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.invoker = invoker;
        this.refreshTokenStore = refreshTokenStore;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
    }

    /**
     * Test constructor — allows injecting a pre-built {@code OfficeExcelAttributes} as a fallback
     * when the attribute store returns null.
     */
    public RestClientProvider(OfficeExcelAttributes officeExcelAttributes,
                              OfficeExcelAttributeStore officeExcelAttributeStore,
                              Invoker invoker, RefreshTokenStore refreshTokenStore,
                              RequestContext requestContext, AuthorizationContext authorizationContext) {
        this.officeExcelAttributes = officeExcelAttributes;
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.invoker = invoker;
        this.refreshTokenStore = refreshTokenStore;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
    }

    public OfficeExcelAttributes getOffice365ExcelAttributes() {
        return officeExcelAttributes;
    }

    /**
     * Creates a REST client for making direct HTTP calls to Microsoft Graph API.
     * Determines the authentication flow (user or admin) based on the request context.
     *
     * @param driveId the ID of the drive to access
     * @return RestClient configured with the appropriate access token
     * @throws MustAuthorizeException if refresh token is not found or has expired
     */
    public RestClient getRestClient(String driveId) {
        Map<String, String> map;
        String key = Constants.EMPTY_STRING;
        try {
            map = requestContext.invokeAsUser() ? handleUserFlow() : handleAdminFlow();

            //find first returns key to be used for refresh token
            Optional<String> value = map.keySet().stream().findFirst();
            if (value.isPresent()) {
                key = value.get();
            }
            String refToken = map.get(key);
            if (refToken == null) {
                NamedValuedField userId = new NamedValuedField("key", "Text", key, new HashMap<>(), new HashMap<>());
                throw new MustAuthorizeException(MessageConstants.REFRESH_TOKEN_NOT_FOUND_RE_AUTHORIZE, List.of(userId));
            }
            return getRestClientOnBehalfOfRefreshToken(refToken, driveId);
        } catch (MustAuthorizeException cause) {
            // Rule 3: MustAuthorizeException must propagate — never wrap or swallow.
            throw cause;
        } catch (Exception cause) {
            NamedValuedField userId = new NamedValuedField("key", "Text", key, new HashMap<>(), new HashMap<>());
            throw new MustAuthorizeException(MessageConstants.REFRESH_TOKEN_EXPIRED_RE_AUTHORIZE, List.of(userId));
        }
    }

    /**
     * Handles user context authentication flow.
     * Retrieves the refresh token for the authorized user account.
     *
     * @return map containing the account ID as key and refresh token as value
     * @throws IOException if refresh token retrieval fails
     */
    private Map<String, String> handleUserFlow() throws IOException {
        Map<String, String> refKey = new HashMap<>();
        String keyValue = authorizationContext.getAuthorizedAccount().getAccountId();
        String refToken = refreshTokenStore.get(keyValue);
        refKey.put(keyValue, refToken);
        return refKey;
    }

    /**
     * Handles admin (service account) authentication flow.
     * Retrieves the refresh token for the admin email configured in attributes.
     *
     * @return map containing the admin email as key and refresh token as value
     * @throws IOException if refresh token retrieval fails
     */
    private Map<String, String> handleAdminFlow() throws IOException {
        Map<String, String> refKey = new HashMap<>();
        OfficeExcelAttributes attrs = resolveAttributes();
        String email = attrs.getEmailId();
        String storeKey = email + Constants.UNDER_SCORE + attrs.getClientId() + Constants.UNDER_SCORE + attrs.getAuthType();
        String refToken = refreshTokenStore.get(storeKey);
        refKey.put(storeKey, refToken);
        return refKey;
    }

    /**
     * Resolves the effective attributes by checking the attribute store first, then falling back
     * to the DI-injected attributes.
     */
    private OfficeExcelAttributes resolveAttributes() {
        String invokerId = invoker.getInvokerId();
        OfficeExcelAttributes stored = officeExcelAttributeStore.load(invokerId);
        if (stored != null) {
            return stored;
        }
        return officeExcelAttributes;
    }

    /**
     * Creates a REST client using a refresh token to acquire an access token.
     * Uses MSAL4J to exchange the refresh token for an access token via OAuth 2.0.
     *
     * @param refToken the OAuth 2.0 refresh token
     * @param driveId the ID of the drive to access
     * @return RestClient configured with the acquired access token
     * @throws MustAuthorizeException if the refresh token is invalid or expired
     */
    public RestClient getRestClientOnBehalfOfRefreshToken(String refToken, String driveId) {


        try {
            OfficeExcelAttributes attrs = resolveAttributes();
            String authority = Constants.PRIVATE.equals(attrs.getAuthType())
                    ? Constants.AUTHORITY + attrs.getTenantId()
                    : Constants.ORG_AUTHORITY;
            String[] scopes = Constants.SCOPE.split(Constants.SPACE);
            Set<String> scopeSet = Arrays.stream(scopes).collect(Collectors.toSet());
            RefreshTokenParameters refreshTokenParameters = RefreshTokenParameters.builder(scopeSet, refToken).build();
            IClientCredential clientCredential = ClientCredentialFactory.createFromSecret(attrs.getClientSecret());
            ConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplication.builder
                    (attrs.getClientId(), clientCredential).authority(authority).build();
            IAuthenticationResult authenticationResult = confidentialClientApplication.acquireToken(refreshTokenParameters).get();
            return new RestClient(authenticationResult.accessToken(), driveId);

        } catch (MustAuthorizeException cause) {
            // Rule 3: MustAuthorizeException must propagate — never wrap or swallow.
            throw cause;
        } catch (InterruptedException cause) {
            LOGGER.debug(cause.getMessage());
            Thread.currentThread().interrupt();
            throw new MustAuthorizeException("The refresh token is no longer valid. Please reauthorize yourself.");
        } catch (Exception cause) {
            throw new MustAuthorizeException("The refresh token is no longer valid. Please reauthorize yourself.");
        }

    }
}
