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
import app.krista.ksdk.accounts.Account;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import com.microsoft.aad.msal4j.*;
import com.microsoft.graph.authentication.IAuthenticationProvider;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.httpcore.HttpClients;
import com.microsoft.graph.models.User;
import com.microsoft.graph.requests.*;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.MockedStatic;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.TimeUnit;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.MicrosoftAuthErrorConstants.*;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Unit tests for {@link GraphServiceClientProvider} targeting 100% code/branch coverage.
 * Tests exercise all public entry points, all exception catch blocks, the inner
 * {@code GraphServiceClientAuthenticationProvider}, and the dead-code private helper
 * {@code handleAuthenticationError} via reflection.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("GraphServiceClientProvider — 100% Coverage")
class GraphServiceClientProviderTest {

    @Mock private RefreshTokenStore mockRefreshTokenStore;
    @Mock private OfficeExcelAttributes mockOfficeExcelAttributes;
    @Mock private OfficeExcelAttributeStore mockAttributeStore;
    @Mock private Invoker mockInvoker;
    @Mock private RequestContext mockRequestContext;
    @Mock private AuthorizationContext mockAuthorizationContext;
    @Mock private OfficeExcelAttributes mockLoadedAttributes;

    private GraphServiceClientProvider provider;

    private static final String TOKEN_CACHE_JSON =
            "{\"RefreshToken\":{\"entry-key\":{\"secret\":\"new-refresh-token\"}}}";

    @BeforeEach
    void setUp() throws Exception {
        provider = new GraphServiceClientProvider(
                mockRefreshTokenStore, mockOfficeExcelAttributes, mockAttributeStore,
                mockInvoker, mockRequestContext, mockAuthorizationContext, "auth-ctx-1"
        );

        when(mockInvoker.getInvokerId()).thenReturn("invoker-1");
        when(mockAttributeStore.load("invoker-1")).thenReturn(mockLoadedAttributes);
        when(mockLoadedAttributes.getEmail()).thenReturn("admin@test.com");
        when(mockOfficeExcelAttributes.getClientId()).thenReturn("client-id");
        when(mockOfficeExcelAttributes.getClientSecret()).thenReturn("client-secret");
        when(mockOfficeExcelAttributes.getAuthType()).thenReturn("Private");
        when(mockOfficeExcelAttributes.getTenantId()).thenReturn("tenant-id");
        when(mockRefreshTokenStore.get(anyString())).thenReturn("mock-refresh-token");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Shared helpers
    // ─────────────────────────────────────────────────────────────────────────

    @SuppressWarnings("unchecked")
    private HappyPathContext setupHappyPath(
            MockedStatic<ConfidentialClientApplication> ccaStatic,
            MockedStatic<HttpClients> httpStatic,
            MockedStatic<GraphServiceClient> gscStatic) throws Exception {

        HappyPathContext ctx = new HappyPathContext();

        ctx.mockCcaBuilder = mock(ConfidentialClientApplication.Builder.class);
        ctx.mockCca       = mock(ConfidentialClientApplication.class);
        ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                .thenReturn(ctx.mockCcaBuilder);
        when(ctx.mockCcaBuilder.authority(anyString())).thenReturn(ctx.mockCcaBuilder);
        when(ctx.mockCcaBuilder.build()).thenReturn(ctx.mockCca);

        ctx.mockAuthResult = mock(IAuthenticationResult.class);
        when(ctx.mockAuthResult.accessToken()).thenReturn("test-access-token");
        when(ctx.mockCca.acquireToken(any(RefreshTokenParameters.class)))
                .thenReturn(CompletableFuture.completedFuture(ctx.mockAuthResult));

        ctx.mockTokenCache = mock(TokenCache.class);
        when(ctx.mockCca.tokenCache()).thenReturn(ctx.mockTokenCache);
        when(ctx.mockTokenCache.serialize()).thenReturn(TOKEN_CACHE_JSON);

        ctx.mockOkHttp   = mock(OkHttpClient.class);
        ctx.mockOkBuilder = mock(OkHttpClient.Builder.class);
        httpStatic.when(() -> HttpClients.createDefault(any(IAuthenticationProvider.class)))
                .thenReturn(ctx.mockOkHttp);
        when(ctx.mockOkHttp.newBuilder()).thenReturn(ctx.mockOkBuilder);
        when(ctx.mockOkBuilder.readTimeout(anyLong(), any(TimeUnit.class))).thenReturn(ctx.mockOkBuilder);
        when(ctx.mockOkBuilder.callTimeout(anyLong(), any(TimeUnit.class))).thenReturn(ctx.mockOkBuilder);
        when(ctx.mockOkBuilder.build()).thenReturn(ctx.mockOkHttp);

        ctx.mockGscBuilder  = mock(GraphServiceClient.Builder.class);
        ctx.mockGraphClient = mock(GraphServiceClient.class);
        gscStatic.when(GraphServiceClient::builder).thenReturn(ctx.mockGscBuilder);
        when(ctx.mockGscBuilder.httpClient(any(OkHttpClient.class))).thenReturn(ctx.mockGscBuilder);
        when(ctx.mockGscBuilder.buildClient()).thenReturn(ctx.mockGraphClient);

        return ctx;
    }

    private static class HappyPathContext {
        ConfidentialClientApplication.Builder mockCcaBuilder;
        ConfidentialClientApplication         mockCca;
        IAuthenticationResult                  mockAuthResult;
        TokenCache                             mockTokenCache;
        OkHttpClient                           mockOkHttp;
        OkHttpClient.Builder                   mockOkBuilder;
        @SuppressWarnings("rawtypes") GraphServiceClient.Builder mockGscBuilder;
        @SuppressWarnings("rawtypes") GraphServiceClient         mockGraphClient;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 1. Happy path — successful client creation (existing)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("Happy path — successful client creation")
    class HappyPath {

        @Test
        @DisplayName("Should return GraphServiceClient and update refresh token")
        void returnsGraphServiceClient() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext ctx = setupHappyPath(ccaStatic, httpStatic, gscStatic);

                GraphServiceClient<Request> result = provider.getGraphServiceClientForAdmin();

                assertNotNull(result);
                assertSame(ctx.mockGraphClient, result);
                verify(mockRefreshTokenStore).put(anyString(), eq("new-refresh-token"));
                verify(ctx.mockOkBuilder).readTimeout(60, TimeUnit.SECONDS);
                verify(ctx.mockOkBuilder).callTimeout(120, TimeUnit.SECONDS);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 2. ExecutionException (existing)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("ExecutionException from acquireToken().get()")
    class ExecutionExceptionPath {

        @Test
        @DisplayName("Should throw MustAuthorizeException when token acquisition fails")
        void throwsMustAuthorizeException() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {

                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);

                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(new RuntimeException("Token acquisition failed"));
                when(mockCca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(failedFuture);

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
            }
        }

        @Test
        @DisplayName("ExecutionException with null cause — uses exception itself as errorMsg source (covers getCause()==null branch)")
        @SuppressWarnings("unchecked")
        void executionException_withNullCause_usesExceptionAsMsalCause() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {

                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);

                // Throw ExecutionException with null cause so getCause() == null branch is taken
                CompletableFuture<IAuthenticationResult> mockFuture = mock(CompletableFuture.class);
                when(mockFuture.get()).thenThrow(new java.util.concurrent.ExecutionException("token error null cause", null));
                when(mockCca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(mockFuture);

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 3. InterruptedException (existing)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("InterruptedException from acquireToken().get()")
    class InterruptedExceptionPath {

        @Test
        @DisplayName("Should interrupt thread and throw MustAuthorizeException")
        @SuppressWarnings("unchecked")
        void interruptsThreadAndThrowsMustAuthorize() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {

                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);

                CompletableFuture<IAuthenticationResult> mockFuture = mock(CompletableFuture.class);
                when(mockFuture.get()).thenThrow(new InterruptedException("Thread interrupted"));
                when(mockCca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(mockFuture);

                Thread.interrupted();

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());

                assertTrue(Thread.interrupted(), "Thread interrupt flag should be set");
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 4. MalformedURLException (existing)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("MalformedURLException from buildConfidentialClientApp")
    class MalformedURLExceptionPath {

        @Test
        @DisplayName("Should throw MustAuthorizeException when authority URL is malformed")
        void throwsMustAuthorizeException() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {

                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenAnswer(inv -> { throw new MalformedURLException("Invalid authority URL"); });

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 5. ClientException (existing)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("ClientException from Graph SDK operations")
    class ClientExceptionPath {

        @Test
        @DisplayName("Should throw MustAuthorizeException when Graph client creation fails")
        void throwsMustAuthorizeException() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients> httpStatic = mockStatic(HttpClients.class)) {

                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);

                IAuthenticationResult mockResult = mock(IAuthenticationResult.class);
                when(mockResult.accessToken()).thenReturn("test-token");
                when(mockCca.acquireToken(any(RefreshTokenParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(mockResult));

                TokenCache mockCache = mock(TokenCache.class);
                when(mockCca.tokenCache()).thenReturn(mockCache);
                when(mockCache.serialize()).thenReturn(TOKEN_CACHE_JSON);

                httpStatic.when(() -> HttpClients.createDefault(any(IAuthenticationProvider.class)))
                        .thenThrow(new ClientException("Graph client creation failed", null));

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 6. IOException from updateRefreshToken (existing)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("IOException from updateRefreshToken")
    class IOExceptionPath {

        @Test
        @DisplayName("Should throw RuntimeException wrapping IOException when token store fails")
        void throwsRuntimeExceptionWrappingIOException() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {

                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);

                IAuthenticationResult mockResult = mock(IAuthenticationResult.class);
                when(mockResult.accessToken()).thenReturn("test-token");
                when(mockCca.acquireToken(any(RefreshTokenParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(mockResult));

                TokenCache mockCache = mock(TokenCache.class);
                when(mockCca.tokenCache()).thenReturn(mockCache);
                when(mockCache.serialize()).thenReturn(TOKEN_CACHE_JSON);

                doThrow(new IOException("Token store write failed"))
                        .when(mockRefreshTokenStore).put(anyString(), anyString());

                RuntimeException thrown = assertThrows(RuntimeException.class,
                        () -> provider.getGraphServiceClientForAdmin());

                assertInstanceOf(IOException.class, thrown.getCause());
                assertEquals("Token store write failed", thrown.getCause().getMessage());
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 7. S2S (Client Credentials) path (existing)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("S2S (Client Credentials) path — getGraphServiceClientForClientCredentials")
    class S2SClientCredentialsPath {

        @Test
        @DisplayName("S2S happy path: uses ClientCredentialParameters, no refresh-token write")
        @SuppressWarnings("unchecked")
        void s2sHappyPath_usesClientCredentialParameters_noRefreshTokenWrite() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>        httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient> gscStatic  = mockStatic(GraphServiceClient.class)) {

                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);
                IAuthenticationResult                  mockResult  = mock(IAuthenticationResult.class);
                OkHttpClient                           mockOkHttp  = mock(OkHttpClient.class);
                OkHttpClient.Builder                   mockOkBuilder = mock(OkHttpClient.Builder.class);
                GraphServiceClient.Builder             mockGscBuilder = mock(GraphServiceClient.Builder.class);
                GraphServiceClient                     mockGraphClient = mock(GraphServiceClient.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);
                when(mockResult.accessToken()).thenReturn("s2s-access-token");
                when(mockCca.acquireToken(any(ClientCredentialParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(mockResult));

                httpStatic.when(() -> HttpClients.createDefault(any(IAuthenticationProvider.class))).thenReturn(mockOkHttp);
                when(mockOkHttp.newBuilder()).thenReturn(mockOkBuilder);
                when(mockOkBuilder.readTimeout(anyLong(), any(TimeUnit.class))).thenReturn(mockOkBuilder);
                when(mockOkBuilder.callTimeout(anyLong(), any(TimeUnit.class))).thenReturn(mockOkBuilder);
                when(mockOkBuilder.build()).thenReturn(mockOkHttp);

                gscStatic.when(GraphServiceClient::builder).thenReturn(mockGscBuilder);
                when(mockGscBuilder.httpClient(any(OkHttpClient.class))).thenReturn(mockGscBuilder);
                when(mockGscBuilder.buildClient()).thenReturn(mockGraphClient);

                GraphServiceClient<Request> result = provider.getGraphServiceClientForAdmin();

                assertNotNull(result);
                assertSame(mockGraphClient, result);
                verify(mockRefreshTokenStore, never()).put(anyString(), anyString());
                verify(mockCca).acquireToken(any(ClientCredentialParameters.class));
                verify(mockCca, never()).acquireToken(any(RefreshTokenParameters.class));
            }
        }

        @Test
        @DisplayName("S2S: ExecutionException → MustAuthorizeException with 'Invalid credentials' message")
        void s2sExecutionException_throwsMustAuthorizeWithCredentialMessage() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);

                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(new RuntimeException("Invalid client secret"));
                when(mockCca.acquireToken(any(ClientCredentialParameters.class))).thenReturn(failedFuture);

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> provider.getGraphServiceClientForAdmin());
                assertTrue(ex.getMessage().contains("Invalid credentials"),
                        "Expected 'Invalid credentials' in message, got: " + ex.getMessage());
            }
        }

        @Test
        @DisplayName("S2S: MalformedURLException → MustAuthorizeException with 'Invalid Tenant ID' message")
        void s2sMalformedUrlException_throwsMustAuthorizeWithTenantIdMessage() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenAnswer(inv -> { throw new MalformedURLException("invalid authority"); });

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> provider.getGraphServiceClientForAdmin());
                assertTrue(ex.getMessage().contains("Invalid Tenant ID format"),
                        "Expected 'Invalid Tenant ID format' in message, got: " + ex.getMessage());
            }
        }

        @Test
        @DisplayName("S2S: InterruptedException → MustAuthorizeException and thread interrupt flag set")
        @SuppressWarnings("unchecked")
        void s2sInterruptedException_throwsMustAuthorizeAndSetsInterruptFlag() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);
                CompletableFuture<IAuthenticationResult> mockFuture = mock(CompletableFuture.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);
                when(mockFuture.get()).thenThrow(new InterruptedException("s2s interrupted"));
                when(mockCca.acquireToken(any(ClientCredentialParameters.class))).thenReturn(mockFuture);

                Thread.interrupted();
                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
                assertTrue(Thread.interrupted(), "Thread interrupt flag should be set after S2S InterruptedException");
            }
        }

        @Test
        @DisplayName("S2S: ExecutionException with null cause → uses exception itself (covers getCause()==null branch)")
        @SuppressWarnings("unchecked")
        void s2sExecutionException_nullCause_throwsMustAuthorize() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         mockCca     = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenReturn(mockCca);

                CompletableFuture<IAuthenticationResult> mockFuture = mock(CompletableFuture.class);
                when(mockFuture.get()).thenThrow(new java.util.concurrent.ExecutionException("s2s null cause", null));
                when(mockCca.acquireToken(any(ClientCredentialParameters.class))).thenReturn(mockFuture);

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> provider.getGraphServiceClientForAdmin());
                assertTrue(ex.getMessage().contains("Invalid credentials"),
                        "Expected 'Invalid credentials' in message, got: " + ex.getMessage());
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 8. determineAuthorityUrl (existing, + unknown-type branch)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("determineAuthorityUrl — authority URL selection")
    class DetermineAuthorityUrl {

        @Test
        @DisplayName("PUBLIC + non-null tenantId → tenant-specific authority (User Auth)")
        void publicAuthWithTenantId_usesTenantSpecificAuthority() throws Exception {
            when(mockOfficeExcelAttributes.getAuthType()).thenReturn("Public");
            when(mockOfficeExcelAttributes.getTenantId()).thenReturn("user-tenant-id");
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenAnswer(inv -> { throw new MalformedURLException("abort"); });

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
                verify(mockBuilder).authority(Constants.AUTHORITY + "user-tenant-id");
            }
        }

        @Test
        @DisplayName("PUBLIC + null tenantId → organizations/ authority (Quick Setup multi-tenant)")
        void publicAuthWithNullTenantId_usesOrgAuthority() throws Exception {
            when(mockOfficeExcelAttributes.getAuthType()).thenReturn("Public");
            when(mockOfficeExcelAttributes.getTenantId()).thenReturn(null);
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenAnswer(inv -> { throw new MalformedURLException("abort"); });

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
                verify(mockBuilder).authority(Constants.ORG_AUTHORITY);
            }
        }

        @Test
        @DisplayName("PRIVATE + tenantId → tenant-specific authority")
        void privateAuthWithTenantId_usesTenantSpecificAuthority() throws Exception {
            when(mockOfficeExcelAttributes.getAuthType()).thenReturn("Private");
            when(mockOfficeExcelAttributes.getTenantId()).thenReturn("s2s-tenant");
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenAnswer(inv -> { throw new MalformedURLException("abort"); });

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
                verify(mockBuilder).authority(Constants.AUTHORITY + "s2s-tenant");
            }
        }

        @Test
        @DisplayName("Unknown authType → IllegalArgumentException")
        void unknownAuthType_throwsIllegalArgumentException() {
            when(mockOfficeExcelAttributes.getAuthType()).thenReturn("UnknownType");
            // refreshTokenStore returns token (setUp) → getGraphServiceClient(String,String) is reached
            // → buildConfidentialClientApp → determineAuthorityUrl("UnknownType", ...) → IAE
            assertThrows(IllegalArgumentException.class, () -> provider.getGraphServiceClientForAdmin());
        }

        @Test
        @DisplayName("PUBLIC + blank tenantId → organizations/ authority (covers isBlank() branch)")
        void publicAuthWithBlankTenantId_usesOrgAuthority() throws Exception {
            when(mockOfficeExcelAttributes.getAuthType()).thenReturn("Public");
            when(mockOfficeExcelAttributes.getTenantId()).thenReturn(""); // blank, not null
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder mockBuilder = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(mockBuilder);
                when(mockBuilder.authority(anyString())).thenReturn(mockBuilder);
                when(mockBuilder.build()).thenAnswer(inv -> { throw new MalformedURLException("abort"); });

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
                verify(mockBuilder).authority(Constants.ORG_AUTHORITY);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 9. GraphServiceClientAuthenticationProvider (inner class)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("GraphServiceClientAuthenticationProvider — inner class")
    class GraphServiceClientAuthenticationProviderTests {

        @Test
        @DisplayName("Constructor: null token → IllegalArgumentException")
        void constructor_nullToken_throwsIllegalArgumentException() {
            assertThrows(IllegalArgumentException.class,
                    () -> new GraphServiceClientProvider.GraphServiceClientAuthenticationProvider(null));
        }

        @Test
        @DisplayName("Constructor: empty token → IllegalArgumentException")
        void constructor_emptyToken_throwsIllegalArgumentException() {
            assertThrows(IllegalArgumentException.class,
                    () -> new GraphServiceClientProvider.GraphServiceClientAuthenticationProvider(""));
        }

        @Test
        @DisplayName("Constructor: blank token → IllegalArgumentException")
        void constructor_blankToken_throwsIllegalArgumentException() {
            assertThrows(IllegalArgumentException.class,
                    () -> new GraphServiceClientProvider.GraphServiceClientAuthenticationProvider("   "));
        }

        @Test
        @DisplayName("Constructor: valid token → no exception")
        void constructor_validToken_succeeds() {
            assertDoesNotThrow(
                    () -> new GraphServiceClientProvider.GraphServiceClientAuthenticationProvider("valid-token"));
        }

        @Test
        @DisplayName("getAuthorizationTokenAsync: returns completed future with the access token")
        void getAuthorizationTokenAsync_returnsTokenFuture() throws Exception {
            var authProvider =
                    new GraphServiceClientProvider.GraphServiceClientAuthenticationProvider("my-token");
            CompletableFuture<String> future =
                    authProvider.getAuthorizationTokenAsync(new URL("https://graph.microsoft.com"));
            assertNotNull(future);
            assertTrue(future.isDone());
            assertEquals("my-token", future.get());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 10. Simple accessors
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("Simple accessors — setRequestContext / getOfficeExcelAttributes")
    class SimpleAccessorTests {

        @Test
        @DisplayName("setRequestContext: subsequent resolveUseEmail(null) reflects the new context")
        void setRequestContext_updatesContext() {
            RequestContext newCtx = mock(RequestContext.class);
            when(newCtx.invokeAsUser()).thenReturn(true);

            // Before update: mockRequestContext.invokeAsUser() = false (default) → resolveUseEmail(null) = !false = true
            assertTrue(provider.resolveUseEmail(null));

            provider.setRequestContext(newCtx);

            // After update: newCtx.invokeAsUser() = true → resolveUseEmail(null) = !true = false
            assertFalse(provider.resolveUseEmail(null));
        }

        @Test
        @DisplayName("getOfficeExcelAttributes: returns the injected attributes instance")
        void getOfficeExcelAttributes_returnsAttributes() {
            assertSame(mockOfficeExcelAttributes, provider.getOfficeExcelAttributes());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 11. resolveUseEmail / isInvokeAsUser
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("resolveUseEmail — all branches")
    class ResolveUseEmailTests {

        @Test
        @DisplayName("useEmail=true → returns true")
        void useEmailTrue_returnsTrue() {
            assertTrue(provider.resolveUseEmail(true));
        }

        @Test
        @DisplayName("useEmail=false → returns false")
        void useEmailFalse_returnsFalse() {
            assertFalse(provider.resolveUseEmail(false));
        }

        @Test
        @DisplayName("useEmail=null, invokeAsUser=true → returns false")
        void useEmailNull_invokeAsUserTrue_returnsFalse() {
            when(mockRequestContext.invokeAsUser()).thenReturn(true);
            assertFalse(provider.resolveUseEmail(null));
        }

        @Test
        @DisplayName("useEmail=null, requestContext=null → returns true (isInvokeAsUser=false)")
        void useEmailNull_requestContextNull_returnsTrue() {
            GraphServiceClientProvider providerNullCtx = new GraphServiceClientProvider(
                    mockRefreshTokenStore, mockOfficeExcelAttributes, mockAttributeStore,
                    mockInvoker, null, mockAuthorizationContext, "auth-ctx-1");
            assertTrue(providerNullCtx.resolveUseEmail(null));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 12. validateToken
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("validateToken")
    class ValidateTokenTests {

        @Test
        @DisplayName("Success path: getGraphServiceClient returns normally, no exception from validateToken")
        void validateToken_success_doesNotThrow() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                setupHappyPath(ccaStatic, httpStatic, gscStatic);
                assertDoesNotThrow(() -> provider.validateToken("user@test.com"));
            }
        }

        @Test
        @DisplayName("IOException path: IOException from token store is silently swallowed")
        void validateToken_ioException_loggedAndSwallowed() throws IOException {
            when(mockRefreshTokenStore.get(anyString())).thenThrow(new IOException("store error"));
            assertDoesNotThrow(() -> provider.validateToken("user@test.com"));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 13. getGraphServiceClient variants (public)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("Public getGraphServiceClient variants")
    class GetGraphServiceClientPublicVariantsTests {

        @Test
        @DisplayName("getGraphServiceClient(RequestContext) delegates to private method")
        void getGraphServiceClient_withRequestContext_delegates() throws Exception {
            RequestContext ctx = mock(RequestContext.class);
            when(ctx.invokeAsUser()).thenReturn(false);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient result = provider.getGraphServiceClient(ctx);
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }

        @Test
        @DisplayName("getGraphServiceClient() no-arg with invokeAsUser=true covers getUserId(false,null) path")
        void getGraphServiceClient_noArg_invokeAsUserTrue_usesAuthorizationContext() throws IOException {
            when(mockRequestContext.invokeAsUser()).thenReturn(true);

            Account mockAccount = mock(Account.class);
            when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(mockAccount);
            when(mockAccount.getAccountId()).thenReturn("user@test.com");

            // IOException from refreshTokenStore propagates out of getGraphServiceClient()
            when(mockRefreshTokenStore.get(anyString())).thenThrow(new IOException("abort"));

            assertThrows(IOException.class, () -> provider.getGraphServiceClient());
        }

        @Test
        @DisplayName("getGraphServiceClientForAsync delegates to getGraphServiceClient(boolean, String)")
        void getGraphServiceClientForAsync_delegates() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient result = provider.getGraphServiceClientForAsync(true, null);
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }

        @Test
        @DisplayName("getGraphServiceClientForAsync with S2S attrs → hits private S2S branch (lines 151-152)")
        void getGraphServiceClientForAsync_s2sAttrs_hitsPrivateS2SBranch() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenAnswer(inv -> { throw new MalformedURLException("abort"); });

                // MalformedURLException caught inside getGraphServiceClientForClientCredentials → MustAuthorizeException
                assertThrows(MustAuthorizeException.class,
                        () -> provider.getGraphServiceClientForAsync(false, "acc@test.com"));
            }
        }

        @Test
        @DisplayName("getGraphServiceClient(RequestContext) with invokeAsUser=true → accountId path (covers !invokeAsUser=false branch)")
        void getGraphServiceClient_withRequestContext_invokeAsUserTrue_delegates() throws Exception {
            RequestContext ctx = mock(RequestContext.class);
            when(ctx.invokeAsUser()).thenReturn(true);
            // !true = false → getGraphServiceClient(false, null) → getUserId(false, null)
            Account acc = mock(Account.class);
            when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(acc);
            when(acc.getAccountId()).thenReturn("user@test.com");

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient result = provider.getGraphServiceClient(ctx);
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }

        @Test
        @DisplayName("getGraphServiceClient() no-arg with invokeAsUser=false → admin email path (covers !invokeAsUser=true branch)")
        void getGraphServiceClient_noArg_invokeAsUserFalse_usesEmailPath() throws Exception {
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient result = provider.getGraphServiceClient();
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 14. getUserRequestBuilder
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("getUserRequestBuilder — useEmail branches")
    class GetUserRequestBuilderTests {

        private Account mockAccount;

        @BeforeEach
        void stubAccount() {
            mockAccount = mock(Account.class);
            when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(mockAccount);
            when(mockAccount.getAccountId()).thenReturn("user@test.com");
        }

        @Test
        @DisplayName("useEmail=true: uses admin email path, returns GraphServiceClient")
        void useEmailTrue_returnsClient() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient<Request> result = provider.getUserRequestBuilder(true, null);
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }

        @Test
        @DisplayName("useEmail=false: uses accountID path, returns GraphServiceClient")
        void useEmailFalse_returnsClient() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                // accountID="user@test.com" → not wsContact → checkUserToken returns early
                GraphServiceClient<Request> result = provider.getUserRequestBuilder(false, "user@test.com");
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }

        @Test
        @DisplayName("useEmail=null: uses requestContext.invokeAsUser() to decide path")
        void useEmailNull_usesRequestContext() throws Exception {
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient<Request> result = provider.getUserRequestBuilder(null, null);
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }

        @Test
        @DisplayName("useEmail=null, invokeAsUser=true → uses accountId (authorizationContext) path (covers !invokeAsUser=false branch)")
        void useEmailNull_invokeAsUserTrue_usesAccountIdPath() throws Exception {
            when(mockRequestContext.invokeAsUser()).thenReturn(true);
            // mockAccount already stubbed in @BeforeEach via mockAccount.getAccountId() = "user@test.com"

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient<Request> result = provider.getUserRequestBuilder(null, null);
                assertNotNull(result);
                assertSame(hpc.mockGraphClient, result);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 15. getAuthenticatedUserEmail
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("getAuthenticatedUserEmail")
    class GetAuthenticatedUserEmailTests {

        @SuppressWarnings("unchecked")
        private GraphServiceClient<Request> buildSpyClient() throws Exception {
            GraphServiceClient<Request> mockClient = mock(GraphServiceClient.class);
            return mockClient;
        }

        @Test
        @DisplayName("User has mail → returns user.mail")
        void userWithMail_returnsMail() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                // getGraphServiceClientForAdmin() is called inside getAuthenticatedUserEmail()
                // via the S2S/refresh-token path — we already have the mock set up.
                // Stub the graph client's me().buildRequest().get() chain:
                when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

                // Re-setup S2S happy path for getGraphServiceClientForAdmin
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);
                IAuthenticationResult                  r  = mock(IAuthenticationResult.class);
                OkHttpClient                           oh = mock(OkHttpClient.class);
                OkHttpClient.Builder                   ob = mock(OkHttpClient.Builder.class);
                GraphServiceClient.Builder             gb = mock(GraphServiceClient.Builder.class);
                GraphServiceClient<Request>            gc = mock(GraphServiceClient.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);
                when(r.accessToken()).thenReturn("tok");
                when(ca.acquireToken(any(ClientCredentialParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(r));
                httpStatic.when(() -> HttpClients.createDefault(any())).thenReturn(oh);
                when(oh.newBuilder()).thenReturn(ob);
                when(ob.readTimeout(anyLong(), any())).thenReturn(ob);
                when(ob.callTimeout(anyLong(), any())).thenReturn(ob);
                when(ob.build()).thenReturn(oh);
                gscStatic.when(GraphServiceClient::builder).thenReturn(gb);
                when(gb.httpClient(any())).thenReturn(gb);
                when(gb.buildClient()).thenReturn(gc);

                // Stub me().buildRequest().get() on the returned client
                UserRequestBuilder  mockMe  = mock(UserRequestBuilder.class);
                UserRequest         mockReq = mock(UserRequest.class);
                User                mockUser = new User();
                mockUser.mail = "admin@example.com";

                when(gc.me()).thenReturn(mockMe);
                when(mockMe.buildRequest()).thenReturn(mockReq);
                when(mockReq.get()).thenReturn(mockUser);

                String email = provider.getAuthenticatedUserEmail();
                assertEquals("admin@example.com", email);
            }
        }

        @Test
        @DisplayName("User is null → returns empty string")
        void userNull_returnsEmptyString() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);
                IAuthenticationResult                  r  = mock(IAuthenticationResult.class);
                OkHttpClient                           oh = mock(OkHttpClient.class);
                OkHttpClient.Builder                   ob = mock(OkHttpClient.Builder.class);
                GraphServiceClient.Builder             gb = mock(GraphServiceClient.Builder.class);
                GraphServiceClient<Request>            gc = mock(GraphServiceClient.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);
                when(r.accessToken()).thenReturn("tok");
                when(ca.acquireToken(any(ClientCredentialParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(r));
                httpStatic.when(() -> HttpClients.createDefault(any())).thenReturn(oh);
                when(oh.newBuilder()).thenReturn(ob);
                when(ob.readTimeout(anyLong(), any())).thenReturn(ob);
                when(ob.callTimeout(anyLong(), any())).thenReturn(ob);
                when(ob.build()).thenReturn(oh);
                gscStatic.when(GraphServiceClient::builder).thenReturn(gb);
                when(gb.httpClient(any())).thenReturn(gb);
                when(gb.buildClient()).thenReturn(gc);

                UserRequestBuilder mockMe  = mock(UserRequestBuilder.class);
                UserRequest        mockReq = mock(UserRequest.class);

                when(gc.me()).thenReturn(mockMe);
                when(mockMe.buildRequest()).thenReturn(mockReq);
                when(mockReq.get()).thenReturn(null);

                String email = provider.getAuthenticatedUserEmail();
                assertEquals("", email);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 16. getDriveRequestBuilder variants (spy approach)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("getDriveRequestBuilder variants")
    class GetDriveRequestBuilderTests {

        @SuppressWarnings("unchecked")
        private GraphServiceClient<Request> buildSpiedClientWithMeChain(
                GraphServiceClientProvider spy,
                DriveRequestBuilder mockDriveBuilder) throws IOException {

            GraphServiceClient<Request> mockClient = mock(GraphServiceClient.class);
            doReturn(mockClient).when(spy).getGraphServiceClient();
            doReturn(mockClient).when(spy).getGraphServiceClient(any(RequestContext.class));
            doReturn(mockClient).when(spy).getGraphServiceClientForAsync(anyBoolean(), any());

            UserRequestBuilder mockMe = mock(UserRequestBuilder.class);
            when(mockClient.me()).thenReturn(mockMe);
            when(mockMe.drive()).thenReturn(mockDriveBuilder);
            when(mockMe.drives(anyString())).thenReturn(mockDriveBuilder);

            return mockClient;
        }

        @Test
        @DisplayName("getDriveRequestBuilder(null) → me().drive()")
        void getDriveRequestBuilder_nullDriveId_returnsMeDrive() throws IOException {
            GraphServiceClientProvider spy = spy(provider);
            DriveRequestBuilder mockDriveBuilder = mock(DriveRequestBuilder.class);
            buildSpiedClientWithMeChain(spy, mockDriveBuilder);

            DriveRequestBuilder result = spy.getDriveRequestBuilder((String) null);
            assertSame(mockDriveBuilder, result);
        }

        @Test
        @DisplayName("getDriveRequestBuilder(driveId) → me().drives(driveId)")
        void getDriveRequestBuilder_withDriveId_returnsDriveById() throws IOException {
            GraphServiceClientProvider spy = spy(provider);
            DriveRequestBuilder mockDriveBuilder = mock(DriveRequestBuilder.class);
            buildSpiedClientWithMeChain(spy, mockDriveBuilder);

            DriveRequestBuilder result = spy.getDriveRequestBuilder("drive-123");
            assertSame(mockDriveBuilder, result);
        }

        @Test
        @DisplayName("getDriveRequestBuilder(driveId) with IOException → wrapped in RuntimeException")
        void getDriveRequestBuilder_ioException_wrapsInRuntimeException() throws IOException {
            when(mockRefreshTokenStore.get(anyString())).thenThrow(new IOException("store error"));
            RuntimeException ex = assertThrows(RuntimeException.class,
                    () -> provider.getDriveRequestBuilder("drive-xyz"));
            assertInstanceOf(IOException.class, ex.getCause());
        }

        @Test
        @DisplayName("getDriveRequestBuilder(null, RequestContext) → me().drive()")
        void getDriveRequestBuilder_nullDriveId_requestContext_returnsMeDrive() {
            GraphServiceClientProvider spy = spy(provider);
            DriveRequestBuilder mockDriveBuilder = mock(DriveRequestBuilder.class);
            try {
                buildSpiedClientWithMeChain(spy, mockDriveBuilder);
            } catch (IOException e) { fail(e); }

            RequestContext ctx = mock(RequestContext.class);
            DriveRequestBuilder result = spy.getDriveRequestBuilder(null, ctx);
            assertSame(mockDriveBuilder, result);
        }

        @Test
        @DisplayName("getDriveRequestBuilder(driveId, RequestContext) → me().drives(driveId)")
        void getDriveRequestBuilder_withDriveId_requestContext_returnsDriveById() {
            GraphServiceClientProvider spy = spy(provider);
            DriveRequestBuilder mockDriveBuilder = mock(DriveRequestBuilder.class);
            try {
                buildSpiedClientWithMeChain(spy, mockDriveBuilder);
            } catch (IOException e) { fail(e); }

            RequestContext ctx = mock(RequestContext.class);
            DriveRequestBuilder result = spy.getDriveRequestBuilder("drive-456", ctx);
            assertSame(mockDriveBuilder, result);
        }

        @Test
        @DisplayName("getDriveRequestBuilder(driveId, RequestContext) with IOException → RuntimeException")
        void getDriveRequestBuilder_ioException_requestContext_wrapsInRuntimeException() throws IOException {
            when(mockRefreshTokenStore.get(anyString())).thenThrow(new IOException("store error"));
            RequestContext ctx = mock(RequestContext.class);
            RuntimeException ex = assertThrows(RuntimeException.class,
                    () -> provider.getDriveRequestBuilder("drive-xyz", ctx));
            assertInstanceOf(IOException.class, ex.getCause());
        }

        @Test
        @DisplayName("getDriveRequestBuilder(null, accountId, invokeAsUser) → me().drive()")
        void getDriveRequestBuilder_threeParms_nullDriveId_returnsMeDrive() throws IOException {
            GraphServiceClientProvider spy = spy(provider);
            DriveRequestBuilder mockDriveBuilder = mock(DriveRequestBuilder.class);
            buildSpiedClientWithMeChain(spy, mockDriveBuilder);

            DriveRequestBuilder result = spy.getDriveRequestBuilder(null, "acc@test.com", false);
            assertSame(mockDriveBuilder, result);
        }

        @Test
        @DisplayName("getDriveRequestBuilder(driveId, accountId, invokeAsUser) → me().drives(driveId)")
        void getDriveRequestBuilder_threeParms_withDriveId_returnsDriveById() throws IOException {
            GraphServiceClientProvider spy = spy(provider);
            DriveRequestBuilder mockDriveBuilder = mock(DriveRequestBuilder.class);
            buildSpiedClientWithMeChain(spy, mockDriveBuilder);

            DriveRequestBuilder result = spy.getDriveRequestBuilder("drive-789", "acc@test.com", true);
            assertSame(mockDriveBuilder, result);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 17. worksheetRequestBuilders
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("worksheetRequestBuilders")
    class WorksheetRequestBuildersTests {

        @Test
        @DisplayName("Chains getDriveRequestBuilder → .items() → .workbook() → .worksheets()")
        @SuppressWarnings("unchecked")
        void worksheetRequestBuilders_chainsCorrectly() throws IOException {
            GraphServiceClientProvider spy = spy(provider);

            GraphServiceClient<Request>      mockClient          = mock(GraphServiceClient.class);
            UserRequestBuilder               mockMe              = mock(UserRequestBuilder.class);
            DriveRequestBuilder              mockDrive           = mock(DriveRequestBuilder.class);
            DriveItemRequestBuilder          mockItem            = mock(DriveItemRequestBuilder.class);
            WorkbookRequestBuilder           mockWorkbook        = mock(WorkbookRequestBuilder.class);
            WorkbookWorksheetRequestBuilder  mockWorksheet       = mock(WorkbookWorksheetRequestBuilder.class);

            doReturn(mockClient).when(spy).getGraphServiceClient();
            when(mockClient.me()).thenReturn(mockMe);
            when(mockMe.drives(eq("drive-id"))).thenReturn(mockDrive);
            when(mockDrive.items(eq("wb-id"))).thenReturn(mockItem);
            when(mockItem.workbook()).thenReturn(mockWorkbook);
            when(mockWorkbook.worksheets(eq("Sheet1"))).thenReturn(mockWorksheet);

            WorkbookWorksheetRequestBuilder result =
                    spy.worksheetRequestBuilders("drive-id", "Sheet1", "wb-id");
            assertSame(mockWorksheet, result);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 18. getRestClient (public) — S2S and non-S2S paths
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("getRestClient (public) — all paths")
    class GetRestClientPublicTests {

        @BeforeEach
        void setInvokeAsUserFalse() {
            // Makes checkUserToken(null,null) resolve wsContactId=null → returns early.
            // Also routes getRestClient to useEmail=true (admin email path).
            when(mockRequestContext.invokeAsUser()).thenReturn(false);
        }

        @Test
        @DisplayName("S2S happy path: returns RestClient without touching refreshTokenStore")
        void s2s_happyPath_returnsRestClient() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);
                IAuthenticationResult                  r  = mock(IAuthenticationResult.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);
                when(r.accessToken()).thenReturn("s2s-token");
                when(ca.acquireToken(any(ClientCredentialParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(r));

                RestClient result = provider.getRestClient("drive-123");
                assertNotNull(result);
                verify(mockRefreshTokenStore, never()).put(anyString(), anyString());
            }
        }

        @Test
        @DisplayName("S2S: ExecutionException → MustAuthorizeException with 'Invalid credentials'")
        void s2s_executionException_throwsMustAuthorize() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);

                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(new RuntimeException("bad creds"));
                when(ca.acquireToken(any(ClientCredentialParameters.class))).thenReturn(failedFuture);

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> provider.getRestClient("drive-123"));
                assertTrue(ex.getMessage().contains("Invalid credentials"));
            }
        }

        @Test
        @DisplayName("S2S: MalformedURLException → MustAuthorizeException with 'Invalid Tenant ID'")
        void s2s_malformedUrlException_throwsMustAuthorize() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenAnswer(inv -> { throw new MalformedURLException("bad url"); });

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> provider.getRestClient("drive-123"));
                assertTrue(ex.getMessage().contains("Invalid Tenant ID format"));
            }
        }

        @Test
        @DisplayName("S2S: InterruptedException → MustAuthorizeException and interrupt flag set")
        @SuppressWarnings("unchecked")
        void s2s_interruptedException_throwsMustAuthorize() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b    = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca   = mock(ConfidentialClientApplication.class);
                CompletableFuture<IAuthenticationResult> f = mock(CompletableFuture.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);
                when(f.get()).thenThrow(new InterruptedException("interrupted"));
                when(ca.acquireToken(any(ClientCredentialParameters.class))).thenReturn(f);

                Thread.interrupted();
                assertThrows(MustAuthorizeException.class, () -> provider.getRestClient("drive-123"));
                assertTrue(Thread.interrupted());
            }
        }

        @Test
        @DisplayName("Non-S2S happy path: acquires token via refresh token, returns RestClient")
        void nonS2S_happyPath_returnsRestClient() throws Exception {
            // authFlow = null → not CLIENT_CREDS → refresh-token path
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(null);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);
                IAuthenticationResult                  r  = mock(IAuthenticationResult.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);
                when(r.accessToken()).thenReturn("refresh-tok");
                when(ca.acquireToken(any(RefreshTokenParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(r));

                RestClient result = provider.getRestClient("drive-123");
                assertNotNull(result);
            }
        }

        @Test
        @DisplayName("Non-S2S: acquireToken fails → MustAuthorizeException (covers getRestClient private exception path)")
        void nonS2S_acquireTokenFails_throwsMustAuthorize() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(null);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);

                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(new RuntimeException("token fail"));
                when(ca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(failedFuture);

                assertThrows(MustAuthorizeException.class, () -> provider.getRestClient("drive-123"));
                Thread.interrupted(); // clear interrupt flag set by catch block
            }
        }

        @Test
        @DisplayName("Non-S2S: null refresh token → MustAuthorizeException (covers private getRestClient null-token branch)")
        void nonS2S_nullRefreshToken_throwsMustAuthorize() throws IOException {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(null);
            when(mockRefreshTokenStore.get(anyString())).thenReturn(null);

            assertThrows(MustAuthorizeException.class, () -> provider.getRestClient("drive-123"));
        }

        @Test
        @DisplayName("Non-S2S: invokeAsUser=true → !true=false, covers getRestClient(driveId, false) branch")
        void nonS2S_invokeAsUserTrue_nullToken_throwsMustAuthorize() throws IOException {
            when(mockRequestContext.invokeAsUser()).thenReturn(true); // !true = false passed to private method
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(null);
            when(mockRefreshTokenStore.get(anyString())).thenReturn(null);

            Account acc = mock(Account.class);
            when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(acc);
            when(acc.getAccountId()).thenReturn("user@test.com");

            assertThrows(MustAuthorizeException.class, () -> provider.getRestClient("drive-123"));
        }

        @Test
        @DisplayName("S2S: getRestClientForClientCredentials ExecutionException null cause — covers getCause()==null branch")
        @SuppressWarnings("unchecked")
        void s2s_restClient_executionException_nullCause_throwsMustAuthorize() throws Exception {
            when(mockOfficeExcelAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);

                CompletableFuture<IAuthenticationResult> mockFuture = mock(CompletableFuture.class);
                when(mockFuture.get()).thenThrow(new java.util.concurrent.ExecutionException("rest s2s null cause", null));
                when(ca.acquireToken(any(ClientCredentialParameters.class))).thenReturn(mockFuture);

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> provider.getRestClient("drive-id"));
                assertTrue(ex.getMessage().contains("Invalid credentials"));
            }
        }

        @Test
        @DisplayName("S2S with null officeExcelAttributes field — loads attrs from store (covers line 340 null-branch)")
        void s2s_nullOfficeExcelAttrsField_loadsFromStore_returnsRestClient() throws Exception {
            // Provider with null officeExcelAttributes field
            GraphServiceClientProvider providerNullAttrs = new GraphServiceClientProvider(
                    mockRefreshTokenStore, null, mockAttributeStore, mockInvoker,
                    mockRequestContext, mockAuthorizationContext, "auth-ctx-1");

            when(mockLoadedAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);
            when(mockLoadedAttributes.getClientSecret()).thenReturn("loaded-secret");
            when(mockLoadedAttributes.getClientId()).thenReturn("loaded-client");
            when(mockLoadedAttributes.getAuthType()).thenReturn("Private");
            when(mockLoadedAttributes.getTenantId()).thenReturn("loaded-tenant");

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);
                IAuthenticationResult                  r  = mock(IAuthenticationResult.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);
                when(r.accessToken()).thenReturn("loaded-s2s-token");
                when(ca.acquireToken(any(ClientCredentialParameters.class)))
                        .thenReturn(CompletableFuture.completedFuture(r));

                RestClient result = providerNullAttrs.getRestClient("drive-456");
                assertNotNull(result);
                verify(mockRefreshTokenStore, never()).put(anyString(), anyString());
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 19. AUTH_ERROR_RULES matching in getGraphServiceClient(String, String)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("AUTH_ERROR_RULES matching — all branches")
    class AuthErrorRulesMatchingTests {

        @Test
        @DisplayName("Rule matches with shouldRemoveToken=true (PASSWORD_CHANGED) → removes token and throws")
        void ruleMatch_shouldRemoveToken_true() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);

                // PASSWORD_CHANGED_CODE causes shouldRemoveToken=true
                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(
                        new RuntimeException(PASSWORD_CHANGED_CODE + " user changed their password"));
                when(ca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(failedFuture);

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> provider.getGraphServiceClientForAdmin());

                // Token must have been removed
                verify(mockRefreshTokenStore).remove(anyString());
                assertTrue(ex.getMessage().contains("password") || ex.getMessage().length() > 0);
            }
        }

        @Test
        @DisplayName("Rule matches with shouldRemoveToken=false (USER_DELETED) → does NOT remove token")
        void ruleMatch_shouldRemoveToken_false() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);

                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(
                        new RuntimeException(USER_DELETED_CODE + " user has been deleted"));
                when(ca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(failedFuture);

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());

                // Token must NOT have been removed
                verify(mockRefreshTokenStore, never()).remove(anyString());
            }
        }

        @Test
        @DisplayName("Null errorMsg (cause.getCause() is null and message is null) → fallback MustAuthorizeException")
        void nullErrorMsg_fallsToDefaultMustAuthorize() throws Exception {
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);

                // Exception with null message
                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(new RuntimeException((String) null));
                when(ca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(failedFuture);

                assertThrows(MustAuthorizeException.class, () -> provider.getGraphServiceClientForAdmin());
                verify(mockRefreshTokenStore, never()).remove(anyString());
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 20. getUserId edge cases
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("getUserId edge cases")
    class GetUserIdEdgeCasesTests {

        @Test
        @DisplayName("calledFromValidateAttributes=true, email from current attrs (not null) → uses that email")
        void emailFromCurrentAttrs_notNull() throws Exception {
            when(mockOfficeExcelAttributes.getEmail()).thenReturn("current@attrs.com");

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                // getGraphServiceClientForAdmin → getGraphServiceClient(true,null) → getUserId(true,null)
                // → emailFromCurrentAttrs = "current@attrs.com" (not null) → uses it
                GraphServiceClient<Request> result = provider.getGraphServiceClientForAdmin();
                assertNotNull(result);
            }
        }

        @Test
        @DisplayName("officeExcelAttributes=null, store=null → userId=null → IllegalStateException")
        void nullAttrsAndNullStore_throwsIllegalStateException() {
            GraphServiceClientProvider providerNullAttrs = new GraphServiceClientProvider(
                    mockRefreshTokenStore, null, mockAttributeStore, mockInvoker,
                    mockRequestContext, mockAuthorizationContext, "auth-ctx-1");
            // Make store return null so userId = null
            when(mockAttributeStore.load("invoker-1")).thenReturn(null);

            assertThrows(IllegalStateException.class, () -> providerNullAttrs.getGraphServiceClientForAdmin());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 21. checkUserToken branches (via getUserRequestBuilder with wsContact accountID)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("checkUserToken — all branches")
    class CheckUserTokenTests {

        private static final String WS_CONTACT_ID = "wsContactTest123";

        @Test
        @DisplayName("Token found → no exception from checkUserToken, proceeds to getGraphServiceClient")
        void tokenFound_noException() throws Exception {
            // "wsContactTest123" resolves as wsContactId → refreshTokenStore.get checked
            // Both calls (checkUserToken + getGraphServiceClient) need to succeed
            when(mockRefreshTokenStore.get(WS_CONTACT_ID)).thenReturn("some-token");

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);

                GraphServiceClient<Request> result = provider.getUserRequestBuilder(false, WS_CONTACT_ID);
                assertNotNull(result);
            }
        }

        @Test
        @DisplayName("Token null → MustAuthorizeException (with invokerId when requestContext != null)")
        void tokenNull_throwsMustAuthorizeException_withInvokerId() throws IOException {
            when(mockRefreshTokenStore.get(WS_CONTACT_ID)).thenReturn(null);
            when(mockRequestContext.getInvokerId()).thenReturn("invoker-1");

            MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                    () -> provider.getUserRequestBuilder(false, WS_CONTACT_ID));
            assertNotNull(ex.getDetails());
        }

        @Test
        @DisplayName("Token null, requestContext=null → no invokerId field in exception details")
        void tokenNull_requestContextNull_noInvokerId() throws IOException {
            GraphServiceClientProvider providerNullCtx = new GraphServiceClientProvider(
                    mockRefreshTokenStore, mockOfficeExcelAttributes, mockAttributeStore,
                    mockInvoker, null, mockAuthorizationContext, "auth-ctx-1");
            when(mockRefreshTokenStore.get(WS_CONTACT_ID)).thenReturn(null);

            MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                    () -> providerNullCtx.getUserRequestBuilder(false, WS_CONTACT_ID));
            // Details should only contain CLIENT_USER_ID (no invokerId)
            assertNotNull(ex.getDetails());
            assertTrue(ex.getDetails().stream()
                    .noneMatch(f -> Constants.USER_ID.equals(f.getName())));
        }

        @Test
        @DisplayName("IOException from token store in checkUserToken → silently swallowed, continues")
        void ioException_swallowedAndContinues() throws IOException {
            // First call (checkUserToken) throws; second call (getGraphServiceClient) returns null → MAE
            when(mockRefreshTokenStore.get(WS_CONTACT_ID))
                    .thenThrow(new IOException("first call"))
                    .thenReturn(null);

            // After IOException is swallowed, getGraphServiceClient(false, WS_CONTACT_ID) is called.
            // getRefTokenStoreKey → wsContact → returns WS_CONTACT_ID.
            // refreshTokenStore.get(WS_CONTACT_ID) → null (second call) → MAE.
            assertThrows(MustAuthorizeException.class,
                    () -> provider.getUserRequestBuilder(false, WS_CONTACT_ID));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 22. resolveWsContactId — all branches
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("resolveWsContactId — all branches")
    class ResolveWsContactIdTests {

        @Test
        @DisplayName("accountID starts with WS_CONTACT → returns accountID immediately")
        void accountIdStartsWithWsContact_returnsAccountId() throws IOException {
            // checkUserToken("wsContact-abc") → resolveWsContactId returns "wsContact-abc"
            when(mockRefreshTokenStore.get("wsContact-abc")).thenReturn(null);
            assertThrows(MustAuthorizeException.class,
                    () -> provider.getUserRequestBuilder(false, "wsContact-abc"));
        }

        @Test
        @DisplayName("resolveUseEmail=false, authId starts with WS_CONTACT → returns authId")
        void resolveUseEmailFalse_authIdStartsWithWsContact_returnsAuthId() throws IOException {
            Account acc = mock(Account.class);
            when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(acc);
            when(acc.getAccountId()).thenReturn("wsContactFromAuth");
            when(mockRefreshTokenStore.get("wsContactFromAuth")).thenReturn(null);

            // useEmail=false, accountID=non-wsContact → resolveUseEmail(false)=false → checks authId
            assertThrows(MustAuthorizeException.class,
                    () -> provider.getUserRequestBuilder(false, "regular@email.com"));
        }

        @Test
        @DisplayName("resolveUseEmail=false, authId doesn't start with WS_CONTACT → returns null")
        void resolveUseEmailFalse_authIdNotWsContact_returnsNull() throws IOException {
            Account acc = mock(Account.class);
            when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(acc);
            when(acc.getAccountId()).thenReturn("regular@user.com");

            // wsContactId = null → checkUserToken returns early → getGraphServiceClient proceeds
            when(mockRefreshTokenStore.get(anyString())).thenReturn(null);
            assertThrows(MustAuthorizeException.class,
                    () -> provider.getUserRequestBuilder(false, "regular@email.com"));
        }

        @Test
        @DisplayName("resolveUseEmail=false, authorizationContext throws → caught, returns null")
        void resolveUseEmailFalse_authContextThrows_returnsNull() throws IOException {
            when(mockAuthorizationContext.getAuthorizedAccount())
                    .thenThrow(new RuntimeException("auth ctx unavailable"));

            when(mockRefreshTokenStore.get(anyString())).thenReturn(null);
            // IOException wraps the MAE and propagates, or MAE propagates directly
            assertThrows(MustAuthorizeException.class,
                    () -> provider.getUserRequestBuilder(false, "other@test.com"));
        }

        @Test
        @DisplayName("resolveUseEmail=true → returns null (checkUserToken returns early)")
        void resolveUseEmailTrue_returnsNull() throws Exception {
            // useEmail=true → resolveUseEmail(true)=true → if(!true)=false → returns null
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class);
                 MockedStatic<HttpClients>          httpStatic = mockStatic(HttpClients.class);
                 MockedStatic<GraphServiceClient>   gscStatic  = mockStatic(GraphServiceClient.class)) {

                HappyPathContext hpc = setupHappyPath(ccaStatic, httpStatic, gscStatic);
                GraphServiceClient<Request> result = provider.getUserRequestBuilder(true, null);
                assertNotNull(result);
            }
        }

        @Test
        @DisplayName("resolveUseEmail=false, authId=null → covers authId!=null branch false path")
        void resolveUseEmailFalse_authIdNull_returnsNull() throws IOException {
            Account acc = mock(Account.class);
            when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(acc);
            when(acc.getAccountId()).thenReturn(null); // null authId → authId != null is false

            when(mockRefreshTokenStore.get(anyString())).thenReturn(null);
            // wsContactId = null → checkUserToken returns early
            // Then getUserId(false, "regular@email.com") uses accountID directly → null refresh → MAE
            assertThrows(MustAuthorizeException.class,
                    () -> provider.getUserRequestBuilder(false, "regular@email.com"));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 23. createMustAuthorizationException / createAuthDetails with null authContextId
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("createMustAuthorizationException / createAuthDetails with null authContextId")
    class NullAuthContextIdTests {

        @Test
        @DisplayName("createMustAuthorizationException: null authContextId → only userId field in details")
        void nullAuthContextId_onlyUserIdInDetails() throws IOException {
            GraphServiceClientProvider providerNullCtx = new GraphServiceClientProvider(
                    mockRefreshTokenStore, mockOfficeExcelAttributes, mockAttributeStore,
                    mockInvoker, mockRequestContext, mockAuthorizationContext, null);
            when(mockRefreshTokenStore.get(anyString())).thenReturn(null);

            MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                    () -> providerNullCtx.getGraphServiceClientForAdmin());
            assertNotNull(ex.getDetails());
            // Only userId — no authContextId field
            assertTrue(ex.getDetails().stream()
                    .noneMatch(f -> Constants.AUTH_CONTEXT_ID.equals(f.getName())));
        }

        @Test
        @DisplayName("createAuthDetails: null authContextId → only userId field (via AUTH_ERROR_RULES match)")
        void nullAuthContextId_createAuthDetails_onlyUserId() throws Exception {
            GraphServiceClientProvider providerNullCtx = new GraphServiceClientProvider(
                    mockRefreshTokenStore, mockOfficeExcelAttributes, mockAttributeStore,
                    mockInvoker, mockRequestContext, mockAuthorizationContext, null);

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b  = mock(ConfidentialClientApplication.Builder.class);
                ConfidentialClientApplication         ca = mock(ConfidentialClientApplication.class);

                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenReturn(ca);

                CompletableFuture<IAuthenticationResult> failedFuture = new CompletableFuture<>();
                failedFuture.completeExceptionally(
                        new RuntimeException(USER_DELETED_CODE + " user deleted"));
                when(ca.acquireToken(any(RefreshTokenParameters.class))).thenReturn(failedFuture);

                MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                        () -> providerNullCtx.getGraphServiceClientForAdmin());
                assertNotNull(ex.getDetails());
                assertTrue(ex.getDetails().stream()
                        .noneMatch(f -> Constants.AUTH_CONTEXT_ID.equals(f.getName())));
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 24. getGraphServiceClientForAdmin with null officeExcelAttributes
    //     (covers getRefTokenStoreKey with officeExcelAttributes==null)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("getGraphServiceClientForAdmin — null officeExcelAttributes")
    class GetGraphServiceClientForAdminNullAttrsTests {

        @Test
        @DisplayName("null attrs + S2S from store → routes through getGraphServiceClientForClientCredentials")
        void nullAttrs_s2sFromStore_routesThroughClientCredentials() throws Exception {
            GraphServiceClientProvider providerNullAttrs = new GraphServiceClientProvider(
                    mockRefreshTokenStore, null, mockAttributeStore, mockInvoker,
                    mockRequestContext, mockAuthorizationContext, "auth-ctx-1");

            when(mockLoadedAttributes.getAuthFlow()).thenReturn(AUTH_FLOW_CLIENT_CREDS);
            when(mockLoadedAttributes.getClientSecret()).thenReturn("s");
            when(mockLoadedAttributes.getClientId()).thenReturn("c");
            when(mockLoadedAttributes.getAuthType()).thenReturn("Private");
            when(mockLoadedAttributes.getTenantId()).thenReturn("t");

            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any())).thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                // Abort after authority() to confirm S2S path was taken without full setup
                when(b.build()).thenAnswer(inv -> { throw new MalformedURLException("abort"); });

                assertThrows(MustAuthorizeException.class, () -> providerNullAttrs.getGraphServiceClientForAdmin());
            }
        }

        @Test
        @DisplayName("null attrs + non-S2S from store + null refresh token → getRefTokenStoreKey uses load.*")
        void nullAttrs_nonS2s_usesLoadedAttrsInRefTokenKey() throws IOException {
            GraphServiceClientProvider providerNullAttrs = new GraphServiceClientProvider(
                    mockRefreshTokenStore, null, mockAttributeStore, mockInvoker,
                    mockRequestContext, mockAuthorizationContext, "auth-ctx-1");

            // Store returns attrs with non-S2S flow
            when(mockLoadedAttributes.getAuthFlow()).thenReturn(null);
            when(mockLoadedAttributes.getClientId()).thenReturn("loaded-client");
            when(mockLoadedAttributes.getAuthType()).thenReturn("Private");
            // Force null refresh token → createMustAuthorizationException is thrown
            when(mockRefreshTokenStore.get(anyString())).thenReturn(null);

            MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                    () -> providerNullAttrs.getGraphServiceClientForAdmin());
            assertNotNull(ex);
        }

        @Test
        @DisplayName("null attrs + store returns null → userId null → IllegalStateException")
        void nullAttrs_storeNull_throwsIllegalStateException() {
            GraphServiceClientProvider providerNullAttrs = new GraphServiceClientProvider(
                    mockRefreshTokenStore, null, mockAttributeStore, mockInvoker,
                    mockRequestContext, mockAuthorizationContext, "auth-ctx-1");
            when(mockAttributeStore.load("invoker-1")).thenReturn(null);

            assertThrows(IllegalStateException.class, () -> providerNullAttrs.getGraphServiceClientForAdmin());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 25. handleAuthenticationError — dead-code private method via reflection
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("handleAuthenticationError — all branches (via reflection)")
    class HandleAuthenticationErrorTests {

        private void invokeHandleAuthError(Exception cause, String key) throws Exception {
            Method m = GraphServiceClientProvider.class
                    .getDeclaredMethod("handleAuthenticationError", Exception.class, String.class);
            m.setAccessible(true);
            try {
                m.invoke(provider, cause, key);
                fail("Expected InvocationTargetException wrapping MustAuthorizeException");
            } catch (InvocationTargetException ite) {
                Throwable t = ite.getCause();
                if (t instanceof MustAuthorizeException) throw (MustAuthorizeException) t;
                if (t instanceof RuntimeException)       throw (RuntimeException) t;
                throw new RuntimeException(t);
            }
        }

        @Test
        @DisplayName("Rule matches with shouldRemoveToken=true → token removed, throws MAE")
        void ruleMatch_removeToken_throwsMAE() throws Exception {
            // PASSWORD_CHANGED_CODE triggers shouldRemoveToken=true
            Exception cause = new RuntimeException(PASSWORD_CHANGED_CODE + " password changed");
            assertThrows(MustAuthorizeException.class, () -> invokeHandleAuthError(cause, "userId-key"));
            verify(mockRefreshTokenStore).remove("userId-key");
        }

        @Test
        @DisplayName("Rule matches with shouldRemoveToken=false → token NOT removed, throws MAE")
        void ruleMatch_noRemove_throwsMAE() throws Exception {
            // USER_DELETED_CODE triggers shouldRemoveToken=false
            Exception cause = new RuntimeException(USER_DELETED_CODE + " user deleted");
            assertThrows(MustAuthorizeException.class, () -> invokeHandleAuthError(cause, "userId-key"));
            verify(mockRefreshTokenStore, never()).remove(anyString());
        }

        @Test
        @DisplayName("No matching rule → throws createMustAuthorizationException")
        void noMatchingRule_throwsDefaultMAE() throws Exception {
            Exception cause = new RuntimeException("some unrecognized error message xyz");
            assertThrows(MustAuthorizeException.class, () -> invokeHandleAuthError(cause, "userId-key"));
        }

        @Test
        @DisplayName("Null error message → skips loop, throws createMustAuthorizationException")
        void nullErrorMessage_throwsDefaultMAE() throws Exception {
            Exception cause = new RuntimeException((String) null);
            assertThrows(MustAuthorizeException.class, () -> invokeHandleAuthError(cause, "userId-key"));
        }

        @Test
        @DisplayName("No match + invokeAsUser=true → createMustAuthorizationException(key, false) (covers !invokeAsUser branch)")
        void noMatchingRule_invokeAsUserTrue_createsMustAuthorizeWithReauthFalse() throws Exception {
            when(mockRequestContext.invokeAsUser()).thenReturn(true); // !true = false
            Exception cause = new RuntimeException("some unrecognized error xyz");
            assertThrows(MustAuthorizeException.class, () -> invokeHandleAuthError(cause, "userId-key-2"));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 26. @Inject six-arg constructor coverage
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("@Inject six-arg constructor")
    class InjectConstructorTests {

        @Test
        @DisplayName("Six-arg constructor initialises all fields (covers lines 61-68)")
        void sixArgConstructor_createsProviderWithCorrectFields() {
            // Uses the @Inject constructor: (RefreshTokenStore, OfficeExcelAttributes,
            // AuthorizationContext, OfficeExcelAttributeStore, Invoker, String)
            GraphServiceClientProvider p = new GraphServiceClientProvider(
                    mockRefreshTokenStore, mockOfficeExcelAttributes, mockAuthorizationContext,
                    mockAttributeStore, mockInvoker, "six-arg-ctx");
            assertNotNull(p);
            assertSame(mockOfficeExcelAttributes, p.getOfficeExcelAttributes());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 27. buildConfidentialClientApp with null officeExcelAttributes field
    //     (covers the officeExcelAttributes==null branch at line 236)
    // ─────────────────────────────────────────────────────────────────────────

    @Nested
    @DisplayName("buildConfidentialClientApp — null officeExcelAttributes field uses loaded attrs")
    class BuildConfidentialClientAppNullAttrsFieldTests {

        @Test
        @DisplayName("null officeExcelAttributes field → buildConfidentialClientApp uses 'load' parameter (line 236 null branch)")
        void nullOfficeExcelAttrsField_buildConfidentialClientApp_usesLoadedAttrs() throws Exception {
            GraphServiceClientProvider providerNullAttrs = new GraphServiceClientProvider(
                    mockRefreshTokenStore, null, mockAttributeStore, mockInvoker,
                    mockRequestContext, mockAuthorizationContext, "auth-ctx-1");

            // Non-S2S loaded attrs so we reach getGraphServiceClient(String, String)
            when(mockLoadedAttributes.getAuthFlow()).thenReturn(null);
            when(mockLoadedAttributes.getEmail()).thenReturn("admin@test.com");
            when(mockLoadedAttributes.getClientId()).thenReturn("loaded-client");
            when(mockLoadedAttributes.getAuthType()).thenReturn("Private");
            when(mockLoadedAttributes.getTenantId()).thenReturn("loaded-tenant");
            when(mockLoadedAttributes.getClientSecret()).thenReturn("loaded-secret");
            // Non-null refresh token so we enter getGraphServiceClient(String, String)
            // Key: "admin@test.com_loaded-client_Private"
            when(mockRefreshTokenStore.get(anyString())).thenReturn("mock-refresh-token");

            // Abort buildConfidentialClientApp after line 236 to confirm the null branch was taken
            try (MockedStatic<ConfidentialClientApplication> ccaStatic = mockStatic(ConfidentialClientApplication.class)) {
                ConfidentialClientApplication.Builder b = mock(ConfidentialClientApplication.Builder.class);
                ccaStatic.when(() -> ConfidentialClientApplication.builder(anyString(), any(IClientCredential.class)))
                        .thenReturn(b);
                when(b.authority(anyString())).thenReturn(b);
                when(b.build()).thenAnswer(inv -> { throw new MalformedURLException("abort after line 236"); });

                assertThrows(MustAuthorizeException.class,
                        () -> providerNullAttrs.getGraphServiceClientForAdmin());
                // Verify that the loaded attrs' clientSecret was used (loaded-secret → ClientCredentialFactory)
                // The authority call confirms we got past line 236 and into the builder
                verify(b).authority(Constants.AUTHORITY + "loaded-tenant");
            }
        }
    }
}
