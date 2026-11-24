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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extensions.util.KeyValueStore;
import app.krista.sdk.msauth.extension.audit.AuditData;
import app.krista.sdk.msauth.extension.audit.AuditRecord;
import app.krista.sdk.msauth.extension.audit.AuditRecorder;
import app.krista.sdk.msauth.extension.service.MsAuthAuditService;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

/**
 * Office 365 Excel audit service — thin delegate around the SDK's {@link MsAuthAuditService}.
 *
 * <p>All audit features (duration tracking, per-request log capture, auth error classification,
 * ring-buffer KV persistence) are provided generically by the SDK.
 * This class wires the SDK service into HK2 for the Excel extension.</p>
 */
@Service
public class Office365ExcelAuditService implements AuditRecorder {

    private final MsAuthAuditService delegate;

    @Inject
    public Office365ExcelAuditService(KeyValueStore keyValueStore) {
        this.delegate = new MsAuthAuditService(keyValueStore, "excel");
    }

    @Override
    public void record(String invokerId, String eventType, String status,
                       String catalogName, String reason, String exceptionClass,
                       Map<String, String> metadata) {
        delegate.record(invokerId, eventType, status, catalogName, reason, exceptionClass, metadata);
    }

    /**
     * Convenience method for recording MustAuthorizeException auth errors.
     */
    public void recordAuthError(String invokerId, String catalogName, MustAuthorizeException cause) {
        delegate.recordAuthError(invokerId, catalogName, cause);
    }

    /** Returns filtered audit data from KV store. */
    public AuditData getData(String invokerId, Long fromTs, Long toTs, String type, String status) {
        return delegate.getData(invokerId, fromTs, toTs, type, status);
    }

    /**
     * Resets auth error counters and removes auth error records.
     * Called when credentials are saved successfully.
     */
    @Override
    public void resetAuthCounters() {
        delegate.resetAuthCounters();
    }

    /** Clears records within [fromTs, toTs] and recalculates the summary. */
    public void clearRecords(String invokerId, Long fromTs, Long toTs) {
        delegate.clearRecords(invokerId, fromTs, toTs);
    }

    /** Starts log capture and timestamps the beginning of a catalog request. */
    public void startRequest() {
        delegate.startRequest();
    }

    /**
     * Records CATALOG_REQUEST_TRIGGERED, then starts log capture for the request.
     *
     * <p>Order matters: TRIGGERED must be recorded BEFORE {@link #startRequest()} because
     * {@code MsAuthAuditService.record()} calls {@code RequestLogCapture.stopCapture()} for
     * any {@code _TRIGGERED} event. Recording TRIGGERED first means stopCapture() drains
     * nothing (no active capture yet), and the subsequent {@link #startRequest()} then opens
     * a fresh capture window for all real work logs. If the order were reversed, the capture
     * started by startRequest() would be immediately killed by record(TRIGGERED), leaving the
     * final SUCCESS/FAILURE record with "No execution logs captured".</p>
     */
    public void startRequest(String invokerId, String catalogName) {
        record(invokerId, "CATALOG_REQUEST_TRIGGERED", "info", catalogName, null, null, null);
        startRequest();
    }

    /** Returns the all-time global counter (never resets). */
    public Map<String, Object> getGlobalCounter() {
        return delegate.getGlobalCounter();
    }

    /** Returns archive index (list of compressed archive metadata). */
    public List<Map<String, Object>> getArchiveIndex() {
        return delegate.getArchiveIndex();
    }

    /** Returns decompressed archived records by key. */
    public List<AuditRecord> getArchivedRecords(String key) {
        return delegate.getArchivedRecords(key);
    }

    /** Deletes a specific archive from KV store. */
    public void deleteArchive(String key) {
        delegate.deleteArchive(key);
    }

    /** Manually archives all current records. */
    public void archiveAll() {
        delegate.archiveAll();
    }
}
