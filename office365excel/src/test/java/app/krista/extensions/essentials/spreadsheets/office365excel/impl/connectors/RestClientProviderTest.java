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
import app.krista.ksdk.accounts.Account;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import com.microsoft.aad.msal4j.IClientCredential;
import com.microsoft.aad.msal4j.IClientSecret;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.MockedStatic;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;
import java.net.MalformedURLException;
import java.util.concurrent.CompletableFuture;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.mockStatic;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

/**
 * Tests for {@link RestClientProvider} targeting 100% branch/line coverage.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class RestClientProviderTest {

    private static final String INVOKER_ID = "invoker-001";
    private static final String DRIVE_ID = "drive-abc";
    private static final String ACCOUNT_ID = "account-xyz";
    private static final String REFRESH_TOKEN = "refresh-token-secret";
    private static final String ACCESS_TOKEN = "access-token-abc123";
    private static final String EMAIL = "admin@example.com";
    private static final String CLIENT_ID = "client-id-001";
    private static final String CLIENT_SECRET = "client-secret-001";
    private static final String TENANT_ID = "tenant-id-001";
    private static final String AUTH_TYPE_PRIVATE = "Private";
    private static final String AUTH_TYPE_PUBLIC = "Public";

    @Mock
    private OfficeExcelAttributes officeExcelAttributes;
    @Mock
    private OfficeExcelAttributeStore officeExcelAttributeStore;
    @Mock
    private Invoker invoker;
    @Mock
    private RefreshTokenStore refreshTokenStore;
    @Mock
    private RequestContext requestContext;
    @Mock
    private AuthorizationContext authorizationContext;

    private RestClientProvider provider;

    @BeforeEach
    void setUp() {
        provider = new RestClientProvider(
                officeExcelAttributes, officeExcelAttributeStore,
                invoker, refreshTokenStore,
                requestContext, authorizationContext);

        when(invoker.getInvokerId()).thenReturn(INVOKER_ID);
    }

    // ==================== getOffice365ExcelAttributes ====================

    @Test
    void getOffice365ExcelAttributes_returnsDiInjectedAttributes() {
        assertSame(officeExcelAttributes, provider.getOffice365ExcelAttributes());
    }

    // ==================== getRestClient — user flow, token found ====================

    @Test
    void getRestClient_userFlow_tokenFound_returnsRestClient() throws Exception {
        when(requestContext.invokeAsUser()).thenReturn(true);

        Account account = mock(Account.class);
        when(account.getAccountId()).thenReturn(ACCOUNT_ID);
        when(authorizationContext.getAuthorizedAccount()).thenReturn(account);
        when(refreshTokenStore.get(ACCOUNT_ID)).thenReturn(REFRESH_TOKEN);

        stubResolveAttributesFromStore(AUTH_TYPE_PUBLIC);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            stubMsalTokenAcquisition(ccfMock, ccaMock);

            RestClient result = provider.getRestClient(DRIVE_ID);

            assertNotNull(result);
        }
    }

    // ==================== getRestClient — admin flow, token found ====================

    @Test
    void getRestClient_adminFlow_tokenFound_returnsRestClient() throws Exception {
        when(requestContext.invokeAsUser()).thenReturn(false);

        stubResolveAttributesFromStore(AUTH_TYPE_PUBLIC);
        String storeKey = EMAIL + Constants.UNDER_SCORE + CLIENT_ID + Constants.UNDER_SCORE + AUTH_TYPE_PUBLIC;
        when(refreshTokenStore.get(storeKey)).thenReturn(REFRESH_TOKEN);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            stubMsalTokenAcquisition(ccfMock, ccaMock);

            RestClient result = provider.getRestClient(DRIVE_ID);

            assertNotNull(result);
        }
    }

    // ==================== getRestClient — admin flow, DI-injected attributes fallback ====================

    @Test
    void getRestClient_adminFlow_storeNull_fallsToDiAttributes() throws Exception {
        when(requestContext.invokeAsUser()).thenReturn(false);

        // Store returns null -> resolveAttributes falls back to DI-injected
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(null);
        when(officeExcelAttributes.getEmailId()).thenReturn(EMAIL);
        when(officeExcelAttributes.getClientId()).thenReturn(CLIENT_ID);
        when(officeExcelAttributes.getAuthType()).thenReturn(AUTH_TYPE_PUBLIC);
        when(officeExcelAttributes.getClientSecret()).thenReturn(CLIENT_SECRET);

        String storeKey = EMAIL + Constants.UNDER_SCORE + CLIENT_ID + Constants.UNDER_SCORE + AUTH_TYPE_PUBLIC;
        when(refreshTokenStore.get(storeKey)).thenReturn(REFRESH_TOKEN);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            stubMsalTokenAcquisition(ccfMock, ccaMock);

            RestClient result = provider.getRestClient(DRIVE_ID);

            assertNotNull(result);
        }
    }

    // ==================== getRestClient — refresh token null ====================

    @Test
    void getRestClient_userFlow_tokenNull_throwsMustAuthorize() throws Exception {
        when(requestContext.invokeAsUser()).thenReturn(true);

        Account account = mock(Account.class);
        when(account.getAccountId()).thenReturn(ACCOUNT_ID);
        when(authorizationContext.getAuthorizedAccount()).thenReturn(account);
        when(refreshTokenStore.get(ACCOUNT_ID)).thenReturn(null);

        // refToken == null -> MustAuthorizeException thrown inside try
        // catch(MustAuthorizeException) re-throws it unchanged (not re-wrapped)
        MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                () -> provider.getRestClient(DRIVE_ID));

        assertEquals(MessageConstants.REFRESH_TOKEN_NOT_FOUND_RE_AUTHORIZE, ex.getMessage());
    }

    // ==================== getRestClient — generic Exception caught ====================

    @Test
    void getRestClient_genericException_throwsMustAuthorizeExpired() throws Exception {
        when(requestContext.invokeAsUser()).thenReturn(true);

        Account account = mock(Account.class);
        when(account.getAccountId()).thenReturn(ACCOUNT_ID);
        when(authorizationContext.getAuthorizedAccount()).thenReturn(account);
        when(refreshTokenStore.get(ACCOUNT_ID)).thenThrow(new IOException("store failure"));

        MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                () -> provider.getRestClient(DRIVE_ID));

        assertEquals(MessageConstants.REFRESH_TOKEN_EXPIRED_RE_AUTHORIZE, ex.getMessage());
    }

    // ==================== resolveAttributes — store returns value ====================

    @Test
    void getRestClient_adminFlow_resolveAttributes_storeReturnsValue() throws Exception {
        when(requestContext.invokeAsUser()).thenReturn(false);

        OfficeExcelAttributes storedAttrs = mock(OfficeExcelAttributes.class);
        when(storedAttrs.getEmailId()).thenReturn(EMAIL);
        when(storedAttrs.getClientId()).thenReturn(CLIENT_ID);
        when(storedAttrs.getAuthType()).thenReturn(AUTH_TYPE_PUBLIC);
        when(storedAttrs.getClientSecret()).thenReturn(CLIENT_SECRET);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(storedAttrs);

        String storeKey = EMAIL + Constants.UNDER_SCORE + CLIENT_ID + Constants.UNDER_SCORE + AUTH_TYPE_PUBLIC;
        when(refreshTokenStore.get(storeKey)).thenReturn(REFRESH_TOKEN);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            stubMsalTokenAcquisition(ccfMock, ccaMock);

            RestClient result = provider.getRestClient(DRIVE_ID);

            assertNotNull(result);
            verify(storedAttrs).getEmailId();
        }
    }

    // ==================== resolveAttributes — store null, falls back to DI ====================

    @Test
    void getRestClient_adminFlow_attrsIsDiInjected() throws Exception {
        when(requestContext.invokeAsUser()).thenReturn(false);

        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(null);
        when(officeExcelAttributes.getEmailId()).thenReturn(EMAIL);
        when(officeExcelAttributes.getClientId()).thenReturn(CLIENT_ID);
        when(officeExcelAttributes.getAuthType()).thenReturn(AUTH_TYPE_PUBLIC);
        when(officeExcelAttributes.getClientSecret()).thenReturn(CLIENT_SECRET);

        String storeKey = EMAIL + Constants.UNDER_SCORE + CLIENT_ID + Constants.UNDER_SCORE + AUTH_TYPE_PUBLIC;
        when(refreshTokenStore.get(storeKey)).thenReturn(REFRESH_TOKEN);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            stubMsalTokenAcquisition(ccfMock, ccaMock);

            RestClient result = provider.getRestClient(DRIVE_ID);

            assertNotNull(result);
            verify(officeExcelAttributes).getEmailId();
        }
    }

    // ==================== getRestClientOnBehalfOfRefreshToken — Private auth type ====================

    @Test
    void getRestClientOnBehalfOfRefreshToken_privateAuthType_usesAuthorityWithTenant() throws MalformedURLException {
        stubResolveAttributesFromStore(AUTH_TYPE_PRIVATE);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            ConfidentialClientApplication.Builder builder = stubMsalTokenAcquisition(ccfMock, ccaMock);

            RestClient result = provider.getRestClientOnBehalfOfRefreshToken(REFRESH_TOKEN, DRIVE_ID);

            assertNotNull(result);
            verify(builder).authority(Constants.AUTHORITY + TENANT_ID);
        }
    }

    // ==================== getRestClientOnBehalfOfRefreshToken — Public auth type ====================

    @Test
    void getRestClientOnBehalfOfRefreshToken_publicAuthType_usesOrgAuthority() throws MalformedURLException {
        stubResolveAttributesFromStore(AUTH_TYPE_PUBLIC);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            ConfidentialClientApplication.Builder builder = stubMsalTokenAcquisition(ccfMock, ccaMock);

            RestClient result = provider.getRestClientOnBehalfOfRefreshToken(REFRESH_TOKEN, DRIVE_ID);

            assertNotNull(result);
            verify(builder).authority(Constants.ORG_AUTHORITY);
        }
    }

    // ==================== getRestClientOnBehalfOfRefreshToken — InterruptedException ====================

    @Test
    void getRestClientOnBehalfOfRefreshToken_interruptedException_throwsMustAuthorize() throws Exception {
        stubResolveAttributesFromStore(AUTH_TYPE_PUBLIC);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            IClientSecret credential = mock(IClientSecret.class);
            ccfMock.when(() -> ClientCredentialFactory.createFromSecret(anyString())).thenReturn(credential);

            ConfidentialClientApplication app = mock(ConfidentialClientApplication.class);
            ConfidentialClientApplication.Builder builder = mock(ConfidentialClientApplication.Builder.class);
            ccaMock.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                    .thenReturn(builder);
            when(builder.authority(anyString())).thenReturn(builder);
            when(builder.build()).thenReturn(app);

            CompletableFuture<IAuthenticationResult> neverCompleteFuture = new CompletableFuture<>();
            when(app.acquireToken(any(com.microsoft.aad.msal4j.RefreshTokenParameters.class)))
                    .thenReturn(neverCompleteFuture);

            // Interrupt thread so .get() throws InterruptedException immediately
            Thread.currentThread().interrupt();

            MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                    () -> provider.getRestClientOnBehalfOfRefreshToken(REFRESH_TOKEN, DRIVE_ID));

            assertEquals("The refresh token is no longer valid. Please reauthorize yourself.", ex.getMessage());
            assertTrue(Thread.currentThread().isInterrupted());

            // Clear interrupt flag for test cleanup
            Thread.interrupted();
        }
    }

    // ==================== getRestClientOnBehalfOfRefreshToken — generic Exception ====================

    @Test
    void getRestClientOnBehalfOfRefreshToken_genericException_throwsMustAuthorize() throws Exception {
        stubResolveAttributesFromStore(AUTH_TYPE_PUBLIC);

        try (MockedStatic<ClientCredentialFactory> ccfMock = mockStatic(ClientCredentialFactory.class);
             MockedStatic<ConfidentialClientApplication> ccaMock = mockStatic(ConfidentialClientApplication.class)) {

            stubMsalTokenAcquisitionThrows(ccfMock, ccaMock, new RuntimeException("MSAL failure"));

            MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                    () -> provider.getRestClientOnBehalfOfRefreshToken(REFRESH_TOKEN, DRIVE_ID));

            assertEquals("The refresh token is no longer valid. Please reauthorize yourself.", ex.getMessage());
        }
    }

    // ==================== Helpers ====================

    private void stubResolveAttributesFromStore(String authType) {
        OfficeExcelAttributes storedAttrs = mock(OfficeExcelAttributes.class);
        when(storedAttrs.getEmailId()).thenReturn(EMAIL);
        when(storedAttrs.getClientId()).thenReturn(CLIENT_ID);
        when(storedAttrs.getClientSecret()).thenReturn(CLIENT_SECRET);
        when(storedAttrs.getAuthType()).thenReturn(authType);
        when(storedAttrs.getTenantId()).thenReturn(TENANT_ID);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(storedAttrs);
    }

    private ConfidentialClientApplication.Builder stubMsalTokenAcquisition(
            MockedStatic<ClientCredentialFactory> ccfMock,
            MockedStatic<ConfidentialClientApplication> ccaMock) throws MalformedURLException {

        IClientSecret credential = mock(IClientSecret.class);
        ccfMock.when(() -> ClientCredentialFactory.createFromSecret(anyString())).thenReturn(credential);

        ConfidentialClientApplication app = mock(ConfidentialClientApplication.class);
        ConfidentialClientApplication.Builder builder = mock(ConfidentialClientApplication.Builder.class);
        ccaMock.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                .thenReturn(builder);
        when(builder.authority(anyString())).thenReturn(builder);
        when(builder.build()).thenReturn(app);

        IAuthenticationResult authResult = mock(IAuthenticationResult.class);
        when(authResult.accessToken()).thenReturn(ACCESS_TOKEN);
        when(app.acquireToken(any(com.microsoft.aad.msal4j.RefreshTokenParameters.class)))
                .thenReturn(CompletableFuture.completedFuture(authResult));

        return builder;
    }

    private void stubMsalTokenAcquisitionThrows(
            MockedStatic<ClientCredentialFactory> ccfMock,
            MockedStatic<ConfidentialClientApplication> ccaMock,
            Exception exception) throws MalformedURLException {

        IClientSecret credential = mock(IClientSecret.class);
        ccfMock.when(() -> ClientCredentialFactory.createFromSecret(anyString())).thenReturn(credential);

        ConfidentialClientApplication app = mock(ConfidentialClientApplication.class);
        ConfidentialClientApplication.Builder builder = mock(ConfidentialClientApplication.Builder.class);
        ccaMock.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                .thenReturn(builder);
        when(builder.authority(anyString())).thenReturn(builder);
        when(builder.build()).thenReturn(app);

        CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
        failedFuture.completeExceptionally(exception);
        when(app.acquireToken(any(com.microsoft.aad.msal4j.RefreshTokenParameters.class)))
                .thenReturn(failedFuture);
    }
}
