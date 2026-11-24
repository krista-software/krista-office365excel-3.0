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

import app.krista.sdk.msauth.config.AuthFlow;
import app.krista.sdk.msauth.extension.annotations.AuditConfig;
import app.krista.sdk.msauth.extension.annotations.AuthScope;
import app.krista.sdk.msauth.extension.annotations.MsAuthExtension;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class Office365ExcelAuthConfigTest {

    @Test
    void testClassCanBeInstantiated() {
        Office365ExcelAuthConfig config = new Office365ExcelAuthConfig();
        assertNotNull(config);
    }

    @Test
    void testMsAuthExtensionAnnotationPresent() {
        MsAuthExtension annotation = Office365ExcelAuthConfig.class.getAnnotation(MsAuthExtension.class);
        assertNotNull(annotation);
    }

    @Test
    void testMsAuthExtensionValues() {
        MsAuthExtension annotation = Office365ExcelAuthConfig.class.getAnnotation(MsAuthExtension.class);
        assertEquals("excel365extn", annotation.extensionType());
        assertEquals("/excel365extn", annotation.callbackPathPrefix());
        assertTrue(annotation.showScopePanel());
        assertTrue(annotation.supportsOwnCredentials());
        assertTrue(annotation.enableAuditDashboard());
    }

    @Test
    void testMsAuthExtensionHeaderValues() {
        MsAuthExtension annotation = Office365ExcelAuthConfig.class.getAnnotation(MsAuthExtension.class);
        assertEquals("Office 365 Excel Authentication", annotation.headerTitle());
        assertEquals("Configure how Krista connects to your Microsoft Office 365 Excel account", annotation.headerSubtitle());
    }

    @Test
    void testMsAuthExtensionSupportedFlows() {
        MsAuthExtension annotation = Office365ExcelAuthConfig.class.getAnnotation(MsAuthExtension.class);
        AuthFlow[] flows = annotation.supportedFlows();
        assertEquals(3, flows.length);
        assertEquals(AuthFlow.AUTH_CODE_PUBLIC, flows[0]);
        assertEquals(AuthFlow.AUTH_CODE_PRIVATE, flows[1]);
        assertEquals(AuthFlow.CLIENT_CREDENTIALS_SECRET, flows[2]);
    }

    @Test
    void testMsAuthExtensionDefaultScopes() {
        MsAuthExtension annotation = Office365ExcelAuthConfig.class.getAnnotation(MsAuthExtension.class);
        String[] scopes = annotation.defaultScopes();
        assertEquals(4, scopes.length);
        assertEquals("openid", scopes[0]);
        assertEquals("offline_access", scopes[1]);
        assertEquals("Files.ReadWrite", scopes[2]);
        assertEquals("User.Read", scopes[3]);
    }

    @Test
    void testAuthScopeAnnotationsPresent() {
        AuthScope[] scopes = Office365ExcelAuthConfig.class.getAnnotationsByType(AuthScope.class);
        assertNotNull(scopes);
        assertEquals(3, scopes.length);
    }

    @Test
    void testFilesReadWriteAllScope() {
        AuthScope[] scopes = Office365ExcelAuthConfig.class.getAnnotationsByType(AuthScope.class);
        AuthScope filesScope = scopes[0];
        assertEquals("Files.ReadWrite.All", filesScope.value());
        assertEquals("Read and write all files", filesScope.description());
        assertTrue(filesScope.required());
        assertTrue(filesScope.catalogRequests().length > 0);
    }

    @Test
    void testSitesReadWriteAllScope() {
        AuthScope[] scopes = Office365ExcelAuthConfig.class.getAnnotationsByType(AuthScope.class);
        AuthScope sitesScope = scopes[1];
        assertEquals("Sites.ReadWrite.All", sitesScope.value());
        assertEquals("Read and write all sites", sitesScope.description());
    }

    @Test
    void testUserReadScope() {
        AuthScope[] scopes = Office365ExcelAuthConfig.class.getAnnotationsByType(AuthScope.class);
        AuthScope userScope = scopes[2];
        assertEquals("User.Read", userScope.value());
        assertEquals("Read user profile", userScope.description());
    }

    @Test
    void testAuditConfigAnnotationPresent() {
        AuditConfig auditConfig = Office365ExcelAuthConfig.class.getAnnotation(AuditConfig.class);
        assertNotNull(auditConfig);
    }

    @Test
    void testAuditConfigValues() {
        AuditConfig auditConfig = Office365ExcelAuthConfig.class.getAnnotation(AuditConfig.class);
        assertEquals("Office 365 Excel Audit Dashboard", auditConfig.dashboardTitle());
        assertTrue(auditConfig.enableCredentialHealth());
        assertTrue(auditConfig.enableCharts());
        assertEquals("excel-audit", auditConfig.csvExportPrefix());
    }

    @Test
    void testSetupGuideUrl() {
        MsAuthExtension annotation = Office365ExcelAuthConfig.class.getAnnotation(MsAuthExtension.class);
        assertEquals("../../../static/docs/#/pages/CreatingOffice365ExcelApp", annotation.setupGuideUrl());
    }
}
