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
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.WorkbookCreateSessionParameterSet;
import com.microsoft.graph.models.WorkbookSessionInfo;
import com.microsoft.graph.options.HeaderOption;
import com.microsoft.graph.options.Option;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.TimeUnit;

/**
 * Manages Microsoft Graph workbook sessions to avoid per-call cold loads that
 * cause {@code MaxRequestDurationExceeded} timeouts on large Excel workbooks.
 * <p>
 * Sessions are created via {@code POST /workbook/createSession} and their ids are
 * reused across subsequent workbook calls via the {@code workbook-session-id}
 * header. Sessions are cached with a short TTL (shorter than Microsoft's ~7 min
 * idle timeout) to maximize reuse while avoiding stale handles.
 */
@Service
public class WorkbookSessionManager {

    private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookSessionManager.class);
    static final String WORKBOOK_SESSION_ID_HEADER = "workbook-session-id";
    private static final long SESSION_TTL_MS = TimeUnit.MINUTES.toMillis(5);
    private static final int DEFAULT_MAX_RETRIES = 3;
    private static final long DEFAULT_INITIAL_BACKOFF_MS = 1000L;

    private final GraphServiceClientProviderFactory providerFactory;
    private GraphServiceClientProvider clientProvider;
    // Package-private for test configuration
    int maxSessionCreateRetries = DEFAULT_MAX_RETRIES;
    long initialBackoffMs = DEFAULT_INITIAL_BACKOFF_MS;
    private final ConcurrentHashMap<String, CachedSession> sessionCache = new ConcurrentHashMap<>();

    @Inject
    public WorkbookSessionManager(GraphServiceClientProviderFactory providerFactory) {
        this.providerFactory = providerFactory;
    }

    WorkbookSessionManager(GraphServiceClientProvider clientProvider) {
        this.providerFactory = null;
        this.clientProvider = clientProvider;
    }

    /** Lazy provider creation — avoids MustAuthorizeException during HK2 construction. */
    private GraphServiceClientProvider getClientProvider() {
        if (clientProvider == null) {
            clientProvider = providerFactory.create();
        }
        return clientProvider;
    }

    /**
     * Returns request options with the {@code workbook-session-id} header set for
     * the given workbook, or an empty list if no session could be created.
     * <p>
     * Callers should always fall back to the returned list as-is; an empty list
     * means requests will go through without a session (same behavior as before).
     *
     * @param driveId         the drive id (null for the user's default drive)
     * @param fileId          the workbook (file) id
     * @param persistChanges  {@code false} for read-only sessions (cheaper/faster);
     *                        {@code true} for write sessions that flush changes
     * @return non-null list containing the session header, or an empty list on failure
     */
    public List<Option> sessionHeaderOptions(String driveId, String fileId, boolean persistChanges) {
        String sessionId = getOrCreateSessionId(driveId, fileId, persistChanges);
        if (sessionId == null || sessionId.isBlank()) {
            return Collections.emptyList();
        }
        return Collections.singletonList(new HeaderOption(WORKBOOK_SESSION_ID_HEADER, sessionId));
    }

    /**
     * Returns a cached session id if still valid, otherwise creates a new one.
     * Session creation is retried with exponential backoff on transient errors
     * (429, 5xx). Returns {@code null} on any non-retryable failure so callers
     * can fall back gracefully.
     */
    public String getOrCreateSessionId(String driveId, String fileId, boolean persistChanges) {
        if (fileId == null || fileId.isBlank()) {
            LOGGER.info("[session] fileId is blank; skipping session creation (drive='{}')", driveId);
            return null;
        }

        String cacheKey = buildCacheKey(driveId, fileId, persistChanges);
        long now = System.currentTimeMillis();

        CachedSession cached = sessionCache.get(cacheKey);
        if (cached != null && cached.expiresAt > now) {
            LOGGER.info("[session] Cache HIT for file='{}' drive='{}' persistChanges={} ttlRemainingMs={}",
                    fileId, driveId, persistChanges, cached.expiresAt - now);
            return cached.sessionId;
        }
        LOGGER.info("[session] Cache MISS for file='{}' drive='{}' persistChanges={} (reason={}); creating new session",
                fileId, driveId, persistChanges, cached == null ? "not-cached" : "expired");

        return createSessionWithRetry(driveId, fileId, persistChanges, cacheKey);
    }

    /**
     * Invalidates the cached session and creates a fresh one. Use this when a
     * previously valid session is rejected by Graph (e.g. expired WAC token).
     *
     * @return new session id, or {@code null} if creation fails
     */
    public String refreshSessionId(String driveId, String fileId, boolean persistChanges) {
        LOGGER.info("[session] Refreshing session for file='{}' drive='{}' persistChanges={}", fileId, driveId, persistChanges);
        invalidate(driveId, fileId, persistChanges);
        return getOrCreateSessionId(driveId, fileId, persistChanges);
    }

    /**
     * Returns request options with a freshly created session, invalidating any
     * cached session first. Use this when retrying after a session rejection.
     */
    public List<Option> refreshSessionHeaderOptions(String driveId, String fileId, boolean persistChanges) {
        String sessionId = refreshSessionId(driveId, fileId, persistChanges);
        if (sessionId == null || sessionId.isBlank()) {
            return Collections.emptyList();
        }
        return Collections.singletonList(new HeaderOption(WORKBOOK_SESSION_ID_HEADER, sessionId));
    }

    private String createSessionWithRetry(String driveId, String fileId, boolean persistChanges, String cacheKey) {
        for (int attempt = 1; attempt <= maxSessionCreateRetries; attempt++) {
            long startNanos = System.nanoTime();
            try {
                WorkbookCreateSessionParameterSet params = WorkbookCreateSessionParameterSet.newBuilder()
                        .withPersistChanges(persistChanges)
                        .build();
                LOGGER.info("[session] params response: {}", params != null ? app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.GSON.toJson(params) : "null");


                WorkbookSessionInfo info = getClientProvider().getDriveRequestBuilder(driveId)
                        .items(fileId)
                        .workbook()
                        .createSession(params)
                        .buildRequest()
                        .post();
                LOGGER.info("[session] Raw createSession response: {}", info != null ? app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.GSON.toJson(info) : "null");

                long elapsedMs = (System.nanoTime() - startNanos) / 1_000_000L;
                if (info == null || info.id == null || info.id.isBlank()) {
                    LOGGER.info("[session] createSession returned empty id for file='{}' drive='{}' elapsedMs={} attempt={}/{}",
                            fileId, driveId, elapsedMs, attempt, maxSessionCreateRetries);
                    if (attempt < maxSessionCreateRetries) {
                        sleepWithBackoff(attempt);
                        continue;
                    }
                    return null;
                }

                long now = System.currentTimeMillis();
                sessionCache.put(cacheKey, new CachedSession(info.id, now + SESSION_TTL_MS));
                LOGGER.info("[session] Created session file='{}' drive='{}' persistChanges={} elapsedMs={} attempt={} idPrefix='{}'",
                        fileId, driveId, persistChanges, elapsedMs, attempt,
                        info.id.substring(0, Math.min(12, info.id.length())));
                return info.id;
            } catch (GraphServiceException cause) {
                long elapsedMs = (System.nanoTime() - startNanos) / 1_000_000L;
                String errorCode = cause.getError() != null && cause.getError().error != null
                        ? cause.getError().error.code : null;

                if (!isRetryableError(cause) || attempt == maxSessionCreateRetries) {
                    LOGGER.info("[session] createSession FAILED file='{}' drive='{}' httpStatus={} errorCode='{}' msg='{}' elapsedMs={} attempt={}/{}",
                            fileId, driveId, cause.getResponseCode(), errorCode, cause.getMessage(), elapsedMs,
                            attempt, maxSessionCreateRetries, cause);
                    return null;
                }
                LOGGER.info("[session] createSession FAILED (will retry) file='{}' drive='{}' httpStatus={} errorCode='{}' elapsedMs={} attempt={}/{}",
                        fileId, driveId, cause.getResponseCode(), errorCode, elapsedMs, attempt, maxSessionCreateRetries);
                sleepWithBackoff(attempt);
            } catch (RuntimeException cause) {
                long elapsedMs = (System.nanoTime() - startNanos) / 1_000_000L;
                LOGGER.info("[session] Unexpected error creating session file='{}' drive='{}' msg='{}' elapsedMs={} attempt={}/{}",
                        fileId, driveId, cause.getMessage(), elapsedMs, attempt, maxSessionCreateRetries, cause);
                if (attempt == maxSessionCreateRetries) {
                    return null;
                }
                sleepWithBackoff(attempt);
            }
        }
        return null;
    }

    private static boolean isRetryableError(GraphServiceException cause) {
        int status = cause.getResponseCode();
        return status == 429 || status >= 500;
    }

    private void sleepWithBackoff(int attempt) {
        long backoffMs = initialBackoffMs * (1L << (attempt - 1));
        if (backoffMs <= 0) {
            return;
        }
        LOGGER.info("[session] Backing off for {}ms before retry attempt {}", backoffMs, attempt + 1);
        try {
            Thread.sleep(backoffMs);
        } catch (InterruptedException ie) {
            Thread.currentThread().interrupt();
        }
    }

    /**
     * Invalidates any cached session for the given workbook. Call this when a
     * previously valid session id is rejected by Graph (e.g. server-side eviction).
     */
    public void invalidate(String driveId, String fileId, boolean persistChanges) {
        sessionCache.remove(buildCacheKey(driveId, fileId, persistChanges));
    }

    private static String buildCacheKey(String driveId, String fileId, boolean persistChanges) {
        return (driveId == null ? "me" : driveId) + "|" + fileId + "|" + persistChanges;
    }

    private static final class CachedSession {
        final String sessionId;
        final long expiresAt;

        CachedSession(String sessionId, long expiresAt) {
            this.sessionId = sessionId;
            this.expiresAt = expiresAt;
        }
    }
}
