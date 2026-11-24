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

import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.sdk.msauth.MicrosoftAuthSdk;
import app.krista.sdk.msauth.extension.audit.AbstractAuditResource;
import app.krista.sdk.msauth.extension.audit.AuditData;
import app.krista.sdk.msauth.extension.audit.AuditRecord;

import javax.inject.Inject;
import javax.ws.rs.Path;
import java.util.*;

/**
 * REST resource for the Office 365 Excel Audit Dashboard.
 *
 * <p>Extends {@link AbstractAuditResource} from the Microsoft Auth SDK, which provides
 * all standard audit endpoints. This class adds credential health checks specific
 * to the Excel extension.</p>
 *
 * <p>Available at: {@code rest/excel365extn/audit/}</p>
 */
@Path("/audit")
public class Office365ExcelAuditResource extends AbstractAuditResource {

    private final Office365ExcelAuditService auditService;
    private final OfficeExcelAttributeStore attributeStore;
    private final RefreshTokenStore refreshTokenStore;

    @Inject
    public Office365ExcelAuditResource(Office365ExcelAuditService auditService,
                                       OfficeExcelAttributeStore attributeStore,
                                       RefreshTokenStore refreshTokenStore,
                                       Invoker invoker) {
        super(invoker.getInvokerId());
        this.auditService = auditService;
        this.attributeStore = attributeStore;
        this.refreshTokenStore = refreshTokenStore;
    }

    @Override
    protected AuditData getAuditData(String invokerId, Long from, Long to, String type, String status) {
        return auditService.getData(invokerId, from, to, type, status);
    }

    @Override
    protected void clearAuditRecords(String invokerId, Long from, Long to) {
        auditService.clearRecords(invokerId, from, to);
    }

    @Override
    protected Map<String, Object> buildHealthResponse() {
        Map<String, Object> health = new HashMap<>();

        // ── Credential configuration ──
        try {
            OfficeExcelAttributes attrs = attributeStore.load(getInvokerId());
            if (attrs != null) {
                health.put("authType", attrs.getAuthType());
                health.put("email", attrs.getEmail());
                health.put("tenantId", mask(attrs.getTenantId()));
                health.put("clientId", mask(attrs.getClientId()));

                // Public client secret expiry from config.properties
                String expiry = null;
                try {
                    java.io.InputStream is = getClass().getClassLoader().getResourceAsStream("config.properties");
                    if (is != null) {
                        java.util.Properties props = new java.util.Properties();
                        props.load(is);
                        expiry = props.getProperty("public.expiryDate");
                        is.close();
                    }
                } catch (Exception ignored) {}
                health.put("clientSecretExpiry", expiry != null ? expiry : "\u2014");

                // Refresh token presence — client credentials (Private) don't use refresh tokens
                if ("Private".equals(attrs.getAuthType())) {
                    health.put("hasRefreshToken", true); // not applicable for client credentials
                } else {
                    // Rule 1: operator keyspace only — no fallback to per-user or bare-email keys.
                    String email = attrs.getEmail();
                    String compositeKey = MicrosoftAuthSdk.buildTokenKey(
                            email, attrs.getClientId(), attrs.getAuthType());
                    String refreshToken = refreshTokenStore.get(compositeKey);
                    health.put("hasRefreshToken", refreshToken != null && !refreshToken.isEmpty());
                }
            } else {
                health.put("authType", "not configured");
                health.put("hasRefreshToken", false);
            }
        } catch (Exception e) {
            health.put("authType", "error");
            health.put("hasRefreshToken", false);
        }

        // ── Auth event counts from audit records ──
        AuditData data = auditService.getData(getInvokerId(), null, null, null, null);
        List<AuditRecord> records = data.records;

        Set<String> authEventKeys = Set.of(
                "TOKEN_REFRESHED", "TOKEN_EXPIRY_DETECTED",
                "CLIENT_SECRET_EXPIRY_DETECTED", "INVALID_CLIENT_SECRET_DETECTED",
                "AUTH_CONFIG_FAILURE", "PASSWORD_CHANGE_DETECTED"
        );

        Map<String, Map<String, Object>> events = new HashMap<>();
        for (String key : authEventKeys) {
            Map<String, Object> ev = new HashMap<>();
            ev.put("count", 0);
            ev.put("lastSeen", null);
            events.put(key, ev);
        }
        for (AuditRecord r : records) {
            if (r.eventType != null && events.containsKey(r.eventType)) {
                Map<String, Object> ev = events.get(r.eventType);
                int cnt = (int) ev.get("count");
                ev.put("count", cnt + 1);
                Long prev = (Long) ev.get("lastSeen");
                if (prev == null || r.timestamp > prev) {
                    ev.put("lastSeen", r.timestamp);
                }
            }
        }
        health.put("events", events);

        return health;
    }

    @Override
    protected List<Map<String, Object>> getArchiveIndex() {
        return auditService.getArchiveIndex();
    }

    @Override
    protected List<AuditRecord> getArchivedRecords(String archiveKey) {
        return auditService.getArchivedRecords(archiveKey);
    }

    @Override
    protected void deleteArchiveEntry(String archiveKey) {
        auditService.deleteArchive(archiveKey);
    }

    @Override
    protected Map<String, Object> getGlobalCounter() {
        return auditService.getGlobalCounter();
    }

    @Override
    protected void archiveAllRecords() {
        auditService.archiveAll();
    }

}
