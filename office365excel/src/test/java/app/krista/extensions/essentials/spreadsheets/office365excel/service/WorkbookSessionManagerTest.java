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

package app.krista.extensions.essentials.spreadsheets.office365excel.service;

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import com.microsoft.graph.http.GraphError;
import com.microsoft.graph.http.GraphErrorResponse;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.WorkbookSessionInfo;
import com.microsoft.graph.options.HeaderOption;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.requests.DriveItemRequestBuilder;
import com.microsoft.graph.requests.DriveRequestBuilder;
import com.microsoft.graph.requests.WorkbookCreateSessionRequest;
import com.microsoft.graph.requests.WorkbookCreateSessionRequestBuilder;
import com.microsoft.graph.requests.WorkbookRequestBuilder;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.lang.reflect.Field;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.times;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

/**
 * Comprehensive test suite for {@link WorkbookSessionManager} targeting 100% branch/line coverage.
 * <p>
 * Covers:
 * <ul>
 *   <li>Both constructors (HK2 {@code @Inject} factory and package-private test constructor)</li>
 *   <li>{@code sessionHeaderOptions}: null / blank / valid sessionId branches</li>
 *   <li>{@code getOrCreateSessionId}: null fileId, blank fileId, cache HIT, cache MISS
 *       (not-cached vs expired), null info, null/blank info.id, success path,
 *       {@link GraphServiceException} with null / non-null error object, generic {@link RuntimeException}</li>
 *   <li>{@code invalidate}</li>
 *   <li>{@code buildCacheKey}: null driveId → {@code "me"}, non-null driveId, persistChanges true/false distinct keys</li>
 *   <li>Nested {@code CachedSession} record-like class constructor</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class WorkbookSessionManagerTest {

    private static final String DRIVE_ID = "drive-123";
    private static final String FILE_ID = "file-abc";
    private static final String SESSION_ID = "session-xyz-0123456789ABCDEF";

    @Mock
    private GraphServiceClientProvider clientProvider;
    @Mock
    private GraphServiceClientProviderFactory providerFactory;
    @Mock
    private DriveRequestBuilder driveRequestBuilder;
    @Mock
    private DriveItemRequestBuilder driveItemRequestBuilder;
    @Mock
    private WorkbookRequestBuilder workbookRequestBuilder;
    @Mock
    private WorkbookCreateSessionRequestBuilder createSessionRequestBuilder;
    @Mock
    private WorkbookCreateSessionRequest createSessionRequest;

    private WorkbookSessionManager manager;

    @BeforeEach
    void setUp() {
        manager = new WorkbookSessionManager(clientProvider);
        // Disable backoff delays in tests.
        manager.initialBackoffMs = 0;
        // Default happy-path chain; individual tests override .post() as needed.
        when(clientProvider.getDriveRequestBuilder(any())).thenReturn(driveRequestBuilder);
        when(driveRequestBuilder.items(any())).thenReturn(driveItemRequestBuilder);
        when(driveItemRequestBuilder.workbook()).thenReturn(workbookRequestBuilder);
        when(workbookRequestBuilder.createSession(any())).thenReturn(createSessionRequestBuilder);
        when(createSessionRequestBuilder.buildRequest()).thenReturn(createSessionRequest);
    }

    // ==================== Constructor Tests ====================

    @Test
    void injectConstructor_usesFactory() {
        // Constructor no longer calls create() eagerly; it stores the factory for lazy init.
        WorkbookSessionManager m = new WorkbookSessionManager(providerFactory);

        assertNotNull(m);
        verify(providerFactory, never()).create();
    }

    @Test
    void packagePrivateConstructor_storesProvider() {
        WorkbookSessionManager m = new WorkbookSessionManager(clientProvider);

        assertNotNull(m);
    }

    // ==================== sessionHeaderOptions Tests ====================

    @Test
    void sessionHeaderOptions_withValidSessionId_returnsHeader() {
        WorkbookSessionInfo info = sessionInfo(SESSION_ID);
        when(createSessionRequest.post()).thenReturn(info);

        List<Option> options = manager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        assertEquals(1, options.size());
        Option opt = options.get(0);
        assertTrue(opt instanceof HeaderOption);
        HeaderOption header = (HeaderOption) opt;
        assertEquals("workbook-session-id", header.getName());
        assertEquals(SESSION_ID, header.getValue());
    }

    @Test
    void sessionHeaderOptions_withNullSessionId_returnsEmpty() {
        // Blank fileId triggers the null-sessionId guard.
        List<Option> options = manager.sessionHeaderOptions(DRIVE_ID, null, false);

        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    @Test
    void sessionHeaderOptions_withBlankSessionId_returnsEmpty() {
        // Force createSession to return an info with a blank id, which propagates as null sessionId.
        when(createSessionRequest.post()).thenReturn(sessionInfo("   "));

        List<Option> options = manager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    @Test
    void sessionHeaderOptions_persistChangesTrue_createsWriteSession() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        List<Option> options = manager.sessionHeaderOptions(DRIVE_ID, FILE_ID, true);

        assertEquals(1, options.size());
    }

    // ==================== getOrCreateSessionId - guard branches ====================

    @Test
    void getOrCreateSessionId_nullFileId_returnsNull() {
        String result = manager.getOrCreateSessionId(DRIVE_ID, null, false);

        assertNull(result);
        verify(clientProvider, never()).getDriveRequestBuilder(any());
    }

    @Test
    void getOrCreateSessionId_blankFileId_returnsNull() {
        String result = manager.getOrCreateSessionId(DRIVE_ID, "   ", false);

        assertNull(result);
        verify(clientProvider, never()).getDriveRequestBuilder(any());
    }

    @Test
    void getOrCreateSessionId_emptyFileId_returnsNull() {
        String result = manager.getOrCreateSessionId(DRIVE_ID, "", false);

        assertNull(result);
        verify(clientProvider, never()).getDriveRequestBuilder(any());
    }

    // ==================== getOrCreateSessionId - cache MISS / success ====================

    @Test
    void getOrCreateSessionId_cacheMiss_createsAndReturnsSession() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(1)).post();
    }

    @Test
    void getOrCreateSessionId_shortId_logsPrefixWithMinBound() {
        // id.length < 12 exercises the Math.min(12, info.id.length()) branch on the logger prefix.
        String shortId = "abc";
        when(createSessionRequest.post()).thenReturn(sessionInfo(shortId));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(shortId, result);
    }

    // ==================== getOrCreateSessionId - cache HIT ====================

    @Test
    void getOrCreateSessionId_cacheHit_returnsCachedWithoutGraphCall() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        // First call populates the cache.
        String first = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        // Second call should hit cache.
        String second = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, first);
        assertEquals(SESSION_ID, second);
        verify(createSessionRequest, times(1)).post(); // Only one actual Graph call.
    }

    @Test
    void getOrCreateSessionId_expiredCache_createsNewSession() throws Exception {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo(SESSION_ID))
                .thenReturn(sessionInfo("session-new-0000"));

        // Prime cache.
        String first = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        assertEquals(SESSION_ID, first);

        // Expire the cached entry by reflecting on sessionCache and rewinding expiresAt.
        expireCachedEntries(manager);

        String second = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        assertEquals("session-new-0000", second);
        verify(createSessionRequest, times(2)).post();
    }

    // ==================== getOrCreateSessionId - null / blank response ====================

    @Test
    void getOrCreateSessionId_nullResponse_returnsNull() {
        when(createSessionRequest.post()).thenReturn(null);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
    }

    @Test
    void getOrCreateSessionId_infoWithNullId_returnsNull() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(null));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
    }

    @Test
    void getOrCreateSessionId_infoWithBlankId_returnsNull() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(""));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
    }

    @Test
    void getOrCreateSessionId_infoWithWhitespaceId_returnsNull() {
        when(createSessionRequest.post()).thenReturn(sessionInfo("   \t  "));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
    }

    // ==================== getOrCreateSessionId - exception branches ====================

    @Test
    void getOrCreateSessionId_graphServiceException_withErrorCode_returnsNull() {
        GraphServiceException ex = mock(GraphServiceException.class);
        GraphErrorResponse errorResponse = new GraphErrorResponse();
        GraphError error = new GraphError();
        error.code = "ItemNotFound";
        errorResponse.error = error;
        when(ex.getError()).thenReturn(errorResponse);
        when(ex.getResponseCode()).thenReturn(404);
        when(ex.getMessage()).thenReturn("Not Found");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
    }

    @Test
    void getOrCreateSessionId_graphServiceException_withNullError_returnsNull() {
        // 500 is retryable, so the mock must throw on all 3 retry attempts.
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(500);
        when(ex.getMessage()).thenReturn("Server error");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(3)).post(); // All 3 attempts failed.
    }

    @Test
    void getOrCreateSessionId_graphServiceException_errorWithNullInnerError_returnsNull() {
        // 500 is retryable, so the mock must throw on all 3 retry attempts.
        GraphServiceException ex = mock(GraphServiceException.class);
        GraphErrorResponse errorResponse = new GraphErrorResponse();
        errorResponse.error = null; // second-half false
        when(ex.getError()).thenReturn(errorResponse);
        when(ex.getResponseCode()).thenReturn(500);
        when(ex.getMessage()).thenReturn("Graph error");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(3)).post(); // All 3 attempts failed.
    }

    @Test
    void getOrCreateSessionId_genericRuntimeException_returnsNull() {
        // RuntimeException is retried, so the mock must throw on all 3 attempts.
        when(createSessionRequest.post())
                .thenThrow(new RuntimeException("boom"))
                .thenThrow(new RuntimeException("boom"))
                .thenThrow(new RuntimeException("boom"));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(3)).post();
    }

    @Test
    void getOrCreateSessionId_retryableException_recoversOnSecondAttempt() {
        // First attempt throws retryable 500, second attempt succeeds.
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(500);
        when(ex.getMessage()).thenReturn("Server error");
        when(createSessionRequest.post())
                .thenThrow(ex)
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void getOrCreateSessionId_nonRetryableException_doesNotRetry() {
        // 403 is non-retryable, should fail immediately.
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(403);
        when(ex.getMessage()).thenReturn("Forbidden");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(1)).post(); // No retry for 403.
    }

    @Test
    void getOrCreateSessionId_exception_doesNotCacheFailedAttempts() {
        // All 3 retry attempts fail, then a fresh call should try again (not use cache).
        when(createSessionRequest.post())
                .thenThrow(new RuntimeException("attempt 1"))
                .thenThrow(new RuntimeException("attempt 2"))
                .thenThrow(new RuntimeException("attempt 3"))
                .thenReturn(sessionInfo(SESSION_ID));

        String first = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        assertNull(first);

        String second = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        assertEquals(SESSION_ID, second);

        verify(createSessionRequest, times(4)).post(); // 3 failed + 1 successful.
    }

    // ==================== Cache key variants (buildCacheKey branches) ====================

    @Test
    void getOrCreateSessionId_nullDriveId_usesMeAsCacheKey() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        String first = manager.getOrCreateSessionId(null, FILE_ID, false);
        String second = manager.getOrCreateSessionId(null, FILE_ID, false);

        assertEquals(SESSION_ID, first);
        assertEquals(SESSION_ID, second);
        verify(createSessionRequest, times(1)).post(); // Cache HIT on 2nd call.
    }

    @Test
    void getOrCreateSessionId_differentDriveIds_areCachedSeparately() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo("session-for-drive1"))
                .thenReturn(sessionInfo("session-for-drive2"));

        String sess1 = manager.getOrCreateSessionId("drive1", FILE_ID, false);
        String sess2 = manager.getOrCreateSessionId("drive2", FILE_ID, false);

        assertEquals("session-for-drive1", sess1);
        assertEquals("session-for-drive2", sess2);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void getOrCreateSessionId_differentPersistChanges_areCachedSeparately() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo("read-session"))
                .thenReturn(sessionInfo("write-session"));

        String readSess = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        String writeSess = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, true);

        assertEquals("read-session", readSess);
        assertEquals("write-session", writeSess);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void getOrCreateSessionId_differentFileIds_areCachedSeparately() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo("session-file1"))
                .thenReturn(sessionInfo("session-file2"));

        String sess1 = manager.getOrCreateSessionId(DRIVE_ID, "file1", false);
        String sess2 = manager.getOrCreateSessionId(DRIVE_ID, "file2", false);

        assertEquals("session-file1", sess1);
        assertEquals("session-file2", sess2);
    }

    // ==================== invalidate Tests ====================

    @Test
    void invalidate_removesCachedEntry() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo(SESSION_ID))
                .thenReturn(sessionInfo("session-after-invalidation"));

        // Prime cache.
        assertEquals(SESSION_ID, manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false));

        // Invalidate.
        manager.invalidate(DRIVE_ID, FILE_ID, false);

        // Next call must re-create the session.
        String refreshed = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        assertEquals("session-after-invalidation", refreshed);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void invalidate_withNonExistentKey_doesNotThrow() {
        // No cache entry exists; invalidate should be a no-op.
        manager.invalidate("never-cached-drive", "never-cached-file", false);
        manager.invalidate(null, "never-cached-file", true);

        // Sanity: next getOrCreateSessionId still performs a single Graph call.
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));
        assertEquals(SESSION_ID, manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false));
    }

    @Test
    void invalidate_onlyAffectsMatchingKey() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo("session-read"))
                .thenReturn(sessionInfo("session-write"));

        // Prime both read and write sessions for the same workbook.
        manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, true);

        // Invalidate only the read session.
        manager.invalidate(DRIVE_ID, FILE_ID, false);

        // Write session still cached (no additional Graph call), read session re-created.
        when(createSessionRequest.post()).thenReturn(sessionInfo("session-read-refreshed"));

        String writeSess = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, true);
        assertEquals("session-write", writeSess); // cached

        String readSess = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        assertEquals("session-read-refreshed", readSess);

        verify(createSessionRequest, times(3)).post();
    }

    // ==================== Header constant ====================

    @Test
    void workbookSessionIdHeaderConstant_isCorrect() {
        assertEquals("workbook-session-id", WorkbookSessionManager.WORKBOOK_SESSION_ID_HEADER);
    }

    @Test
    void sessionHeaderOptions_usesCorrectHeaderName() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        List<Option> options = manager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        HeaderOption header = (HeaderOption) options.get(0);
        assertEquals(WorkbookSessionManager.WORKBOOK_SESSION_ID_HEADER, header.getName());
    }

    // ==================== Graph client chain verification ====================

    @Test
    void getOrCreateSessionId_usesCorrectGraphChain() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, true);

        verify(clientProvider).getDriveRequestBuilder(DRIVE_ID);
        verify(driveRequestBuilder).items(FILE_ID);
        verify(driveItemRequestBuilder).workbook();
        verify(workbookRequestBuilder).createSession(any());
        verify(createSessionRequestBuilder).buildRequest();
        verify(createSessionRequest).post();
    }

    @Test
    void getOrCreateSessionId_nullDriveId_forwardedToProvider() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        manager.getOrCreateSessionId(null, FILE_ID, false);

        verify(clientProvider).getDriveRequestBuilder(null);
    }

    // ==================== Sanity: sessionHeaderOptions empty list is immutable-safe ====================

    @Test
    void sessionHeaderOptions_emptyResult_isNotNull() {
        // Blank fileId → guard → null sessionId → empty list.
        List<Option> options = manager.sessionHeaderOptions(DRIVE_ID, "", true);

        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    @Test
    void sessionHeaderOptions_onGraphException_returnsEmpty() {
        // RuntimeException retries 3 times, all must fail.
        when(createSessionRequest.post())
                .thenThrow(new RuntimeException("transient graph error"))
                .thenThrow(new RuntimeException("transient graph error"))
                .thenThrow(new RuntimeException("transient graph error"));

        List<Option> options = manager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    // ==================== refreshSessionId Tests ====================

    @Test
    void refreshSessionId_invalidatesThenCreatesNewSession() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo(SESSION_ID))
                .thenReturn(sessionInfo("session-refreshed"));

        // Prime cache.
        String first = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
        assertEquals(SESSION_ID, first);

        // Refresh should invalidate + create new.
        String refreshed = manager.refreshSessionId(DRIVE_ID, FILE_ID, false);
        assertEquals("session-refreshed", refreshed);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void refreshSessionId_nullFileId_returnsNull() {
        String result = manager.refreshSessionId(DRIVE_ID, null, false);
        assertNull(result);
    }

    @Test
    void refreshSessionId_creationFails_returnsNull() {
        when(createSessionRequest.post())
                .thenThrow(new RuntimeException("fail"))
                .thenThrow(new RuntimeException("fail"))
                .thenThrow(new RuntimeException("fail"));

        String result = manager.refreshSessionId(DRIVE_ID, FILE_ID, false);
        assertNull(result);
    }

    @Test
    void refreshSessionId_withNullDriveId_works() {
        when(createSessionRequest.post()).thenReturn(sessionInfo("session-null-drive"));

        String result = manager.refreshSessionId(null, FILE_ID, false);
        assertEquals("session-null-drive", result);
    }

    @Test
    void refreshSessionId_persistChangesTrue_works() {
        when(createSessionRequest.post()).thenReturn(sessionInfo("write-session-refreshed"));

        String result = manager.refreshSessionId(DRIVE_ID, FILE_ID, true);
        assertEquals("write-session-refreshed", result);
    }

    // ==================== refreshSessionHeaderOptions Tests ====================

    @Test
    void refreshSessionHeaderOptions_validSession_returnsHeader() {
        when(createSessionRequest.post()).thenReturn(sessionInfo(SESSION_ID));

        List<Option> options = manager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        assertEquals(1, options.size());
        HeaderOption header = (HeaderOption) options.get(0);
        assertEquals("workbook-session-id", header.getName());
        assertEquals(SESSION_ID, header.getValue());
    }

    @Test
    void refreshSessionHeaderOptions_nullSessionId_returnsEmpty() {
        // Null fileId → null sessionId.
        List<Option> options = manager.refreshSessionHeaderOptions(DRIVE_ID, null, false);

        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    @Test
    void refreshSessionHeaderOptions_blankSessionId_returnsEmpty() {
        when(createSessionRequest.post()).thenReturn(sessionInfo("   "));

        List<Option> options = manager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    @Test
    void refreshSessionHeaderOptions_creationFails_returnsEmpty() {
        when(createSessionRequest.post())
                .thenThrow(new RuntimeException("fail"))
                .thenThrow(new RuntimeException("fail"))
                .thenThrow(new RuntimeException("fail"));

        List<Option> options = manager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    @Test
    void refreshSessionHeaderOptions_invalidatesPreviousSession() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo("old-session"))
                .thenReturn(sessionInfo("new-session"));

        // Prime cache.
        manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        // Refresh should create a new session, not return the cached one.
        List<Option> options = manager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false);

        assertEquals(1, options.size());
        HeaderOption header = (HeaderOption) options.get(0);
        assertEquals("new-session", header.getValue());
        verify(createSessionRequest, times(2)).post();
    }

    // ==================== createSessionWithRetry - additional branch coverage ====================

    @Test
    void createSession_retryable429_retriesAndSucceeds() {
        GraphServiceException ex429 = mock(GraphServiceException.class);
        when(ex429.getError()).thenReturn(null);
        when(ex429.getResponseCode()).thenReturn(429);
        when(ex429.getMessage()).thenReturn("Too Many Requests");

        when(createSessionRequest.post())
                .thenThrow(ex429)
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void createSession_retryable429_allAttemptsFail_returnsNull() {
        GraphServiceException ex429 = mock(GraphServiceException.class);
        when(ex429.getError()).thenReturn(null);
        when(ex429.getResponseCode()).thenReturn(429);
        when(ex429.getMessage()).thenReturn("Too Many Requests");
        when(createSessionRequest.post()).thenThrow(ex429);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(3)).post();
    }

    @Test
    void createSession_retryable503_retriesAndSucceeds() {
        GraphServiceException ex503 = mock(GraphServiceException.class);
        when(ex503.getError()).thenReturn(null);
        when(ex503.getResponseCode()).thenReturn(503);
        when(ex503.getMessage()).thenReturn("Service Unavailable");

        when(createSessionRequest.post())
                .thenThrow(ex503)
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void createSession_retryableGraphException_withErrorCode_retriesAndSucceeds() {
        GraphServiceException ex = mock(GraphServiceException.class);
        GraphErrorResponse errorResponse = new GraphErrorResponse();
        GraphError error = new GraphError();
        error.code = "ServiceUnavailable";
        errorResponse.error = error;
        when(ex.getError()).thenReturn(errorResponse);
        when(ex.getResponseCode()).thenReturn(503);
        when(ex.getMessage()).thenReturn("Service Unavailable");

        when(createSessionRequest.post())
                .thenThrow(ex)
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void createSession_runtimeException_recoversOnSecondAttempt() {
        when(createSessionRequest.post())
                .thenThrow(new RuntimeException("transient"))
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void createSession_nullResponse_retriesAndEventuallyReturnsNull() {
        when(createSessionRequest.post())
                .thenReturn(null)
                .thenReturn(null)
                .thenReturn(null);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(3)).post();
    }

    @Test
    void createSession_nullResponse_recoversOnSecondAttempt() {
        when(createSessionRequest.post())
                .thenReturn(null)
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void createSession_blankIdResponse_retriesAndRecovers() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo(""))
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void createSession_nullIdResponse_retriesAndRecovers() {
        when(createSessionRequest.post())
                .thenReturn(sessionInfo(null))
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    // ==================== sleepWithBackoff Tests ====================

    @Test
    void sleepWithBackoff_zeroBackoff_skipsSleep() {
        manager.initialBackoffMs = 0;
        // Should return immediately without sleeping — no exception or hang.
        manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);
    }

    @Test
    void sleepWithBackoff_positiveBackoff_sleeps() {
        // Use 1ms backoff to verify the sleep path executes without hanging.
        manager.initialBackoffMs = 1;

        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(500);
        when(ex.getMessage()).thenReturn("Server error");

        when(createSessionRequest.post())
                .thenThrow(ex)
                .thenReturn(sessionInfo(SESSION_ID));

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertEquals(SESSION_ID, result);
        verify(createSessionRequest, times(2)).post();
    }

    @Test
    void sleepWithBackoff_interruptedDuringSleep_restoresInterruptFlag() throws Exception {
        // Use a short but non-zero backoff so Thread.sleep is actually called.
        manager.initialBackoffMs = 50;

        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(500);
        when(ex.getMessage()).thenReturn("Server error");

        // All 3 attempts fail — sleep will be called between them.
        when(createSessionRequest.post()).thenThrow(ex);

        Thread testThread = Thread.currentThread();

        // Interrupt the thread after a short delay to trigger InterruptedException in sleep.
        Thread interruptor = new Thread(() -> {
            try {
                Thread.sleep(10);
            } catch (InterruptedException ignored) {
            }
            testThread.interrupt();
        });
        interruptor.start();

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        // The interrupt flag should have been restored by sleepWithBackoff.
        assertTrue(Thread.interrupted(), "Interrupt flag should be set after InterruptedException");

        interruptor.join(1000);
    }

    // ==================== isRetryableError - exhaustive status codes ====================

    @Test
    void createSession_status400_notRetryable_failsImmediately() {
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(400);
        when(ex.getMessage()).thenReturn("Bad Request");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(1)).post();
    }

    @Test
    void createSession_status401_notRetryable_failsImmediately() {
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(401);
        when(ex.getMessage()).thenReturn("Unauthorized");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(1)).post();
    }

    @Test
    void createSession_status499_notRetryable_failsImmediately() {
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(499);
        when(ex.getMessage()).thenReturn("Client error");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(1)).post();
    }

    @Test
    void createSession_status500_retryable_retries3Times() {
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(500);
        when(ex.getMessage()).thenReturn("Internal Server Error");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(3)).post();
    }

    @Test
    void createSession_status502_retryable_retries3Times() {
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getError()).thenReturn(null);
        when(ex.getResponseCode()).thenReturn(502);
        when(ex.getMessage()).thenReturn("Bad Gateway");
        when(createSessionRequest.post()).thenThrow(ex);

        String result = manager.getOrCreateSessionId(DRIVE_ID, FILE_ID, false);

        assertNull(result);
        verify(createSessionRequest, times(3)).post();
    }

    // ==================== Helpers ====================

    private static WorkbookSessionInfo sessionInfo(String id) {
        WorkbookSessionInfo info = new WorkbookSessionInfo();
        info.id = id;
        info.persistChanges = Boolean.FALSE;
        return info;
    }

    /**
     * Replaces every cached entry with a freshly-constructed {@code CachedSession} whose
     * {@code expiresAt} is set to epoch, so the next {@code getOrCreateSessionId} call
     * treats it as expired and takes the cache-MISS-due-to-expiry branch.
     * <p>
     * We rebuild the entry rather than mutating its final {@code expiresAt} field because
     * Java 17 disallows reflective writes to {@code final} instance fields even with
     * {@code setAccessible(true)}.
     */
    @SuppressWarnings("unchecked")
    private static void expireCachedEntries(WorkbookSessionManager target) throws Exception {
        Field cacheField = WorkbookSessionManager.class.getDeclaredField("sessionCache");
        cacheField.setAccessible(true);
        ConcurrentHashMap<String, Object> cache = (ConcurrentHashMap<String, Object>) cacheField.get(target);
        assertFalse(cache.isEmpty(), "precondition: cache should contain at least one entry");

        Class<?> cachedSessionClass = Class.forName(
                WorkbookSessionManager.class.getName() + "$CachedSession");
        java.lang.reflect.Constructor<?> ctor =
                cachedSessionClass.getDeclaredConstructor(String.class, long.class);
        ctor.setAccessible(true);

        for (java.util.Map.Entry<String, Object> entry : cache.entrySet()) {
            Field sessionIdField = cachedSessionClass.getDeclaredField("sessionId");
            sessionIdField.setAccessible(true);
            String existingId = (String) sessionIdField.get(entry.getValue());
            // Rebuild with expiresAt=0 so the entry is guaranteed to be treated as expired.
            Object expiredEntry = ctor.newInstance(existingId, 0L);
            cache.put(entry.getKey(), expiredEntry);
        }
    }
}
