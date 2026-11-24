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

import app.krista.extension.authorization.RequestAuthenticator;
import app.krista.extension.executor.Invoker;
import app.krista.extension.impl.anno.*;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;

import javax.inject.Inject;
import java.io.IOException;
import java.util.Map;

@Java(version = Java.Version.JAVA_21)
@ChangeLog(file = "resources/docs/pages/release-notes.md")
@StaticResource(path = "docs", file = "docs")
@Containerize(baseImageVersion = "3.5.9")
@Extension(version = "4.0.4", implementingDomainIds = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24", name = "Office 365 Excel - 3.0", implementationModel = Extension.ImplementationModel.Krista_4_O)
public final class Office365ExcelExtension {

    private final GraphServiceClientProviderFactory clientProviderFactory;
    private final OfficeExcelAttributeStore officeExcelAttributeStore;
    private final OfficeExcelRequestAuthenticator requestAuthenticator;
    private final String routingId;
    private final String invokerId;
    private final Office365ExcelAuditService auditService;

    @Inject
    public Office365ExcelExtension(Invoker invoker,
                                   GraphServiceClientProviderFactory clientProviderFactory,
                                   OfficeExcelAttributeStore officeExcelAttributeStore,
                                   Office365ExcelAuditService auditService) {
        this.clientProviderFactory = clientProviderFactory;
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.auditService = auditService;
        this.invokerId = invoker.getInvokerId();
        this.routingId = invoker.getRoutingInfo().getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE);
        this.requestAuthenticator = new OfficeExcelRequestAuthenticator(
                officeExcelAttributeStore,
                this.invokerId,
                this.routingId);
    }

    /**
     * Returns the SDK-managed OAuth request authenticator.
     * Replaces the old ScribeJava-based authenticator.
     */
    @InvokerRequest(InvokerRequest.Type.AUTHENTICATOR)
    public RequestAuthenticator getRequestAuthenticator() {
        return this.requestAuthenticator;
    }

    /**
     * Defines custom tabs for the extension UI.
     * "Authentication" now serves the pre-built SDK auth tab HTML.
     * "Excel Audit Dashboard" serves the pre-built SDK audit dashboard HTML.
     */
    @InvokerRequest(InvokerRequest.Type.CUSTOM_TABS)
    public Map<String, String> customTabs() {
        return Map.of(
                "Authentication", "rest/excel365extn/authentication/",
                "Documentation", "static/docs",
                "Excel Audit Dashboard", "rest/excel365extn/audit/",
                "AI Assistant", "rest/excel365extn/mcp-setup"
        );
    }

    @InvokerRequest(InvokerRequest.Type.VALIDATE_ATTRIBUTES)
    public void validate(Map<String, Object> attributes) throws IOException {
        OfficeExcelAttributes excelAttributes = OfficeExcelAttributes.create(routingId, attributes);
        testConnection(excelAttributes);
    }

    /**
     * Tests connection to Office365 Excel.
     * Directs users to the Authentication tab for connection testing.
     */
    @InvokerRequest(InvokerRequest.Type.TEST_CONNECTION)
    public void testConnection() {
        throw new IllegalArgumentException("Please use Authentication tab for test connection.");
    }

    @InvokerRequest(InvokerRequest.Type.INVOKER_UPDATED)
    public void attributesUpdated(Map<String, Object> oldAttributes, Map<String, Object> newAttributes) {
        OfficeExcelAttributes updated = OfficeExcelAttributes.create(routingId, newAttributes);
        officeExcelAttributeStore.save(updated, invokerId);
    }

    private void testConnection(OfficeExcelAttributes excelAttributes) throws IOException {
        String email = excelAttributes.getEmailId();
        if (email != null && !email.isBlank()) {
            clientProviderFactory.create(excelAttributes)
                    .getGraphServiceClientForAdmin()
                    .me().buildRequest().get();
        }
    }
}
