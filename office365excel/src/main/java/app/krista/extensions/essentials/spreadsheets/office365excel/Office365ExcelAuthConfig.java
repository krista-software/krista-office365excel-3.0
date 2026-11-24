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
import app.krista.sdk.msauth.extension.annotations.*;

// ── Module 1: Auth SDK Core ──
@MsAuthExtension(
        extensionType = "excel365extn",
        callbackPathPrefix = "/excel365extn",
        supportedFlows = {AuthFlow.AUTH_CODE_PUBLIC, AuthFlow.AUTH_CODE_PRIVATE, AuthFlow.CLIENT_CREDENTIALS_SECRET},
        defaultScopes = {"openid", "offline_access", "Files.ReadWrite", "User.Read"},
        // ── Module 2: Auth Tab UI ──
        headerTitle = "Office 365 Excel Authentication",
        headerSubtitle = "Configure how Krista connects to your Microsoft Office 365 Excel account",
        showScopePanel = true,
        supportsOwnCredentials = true,
        enableAuditDashboard = true,
        setupGuideUrl = "../../../static/docs/#/pages/CreatingOffice365ExcelApp"
)
// ── Module 2: Scopes ──
@AuthScope(value = "Files.ReadWrite.All", description = "Read and write all files", required = true,
        catalogRequests = {"Get All Excel Files For Name", "Get Excel File Information", "Get Sheets Information",
                "Add Rows", "Add Rows Async", "Add Row With Index", "Add Row With Index Async",
                "Update Row With Index", "Add Data To Cell", "Add Multiple Cell Data",
                "Add Multiple Cell Data Async", "Delete Sheet", "Delete Row With Index",
                "Delete Column With Index", "Clear Sheet For Range", "Find Value",
                "Get Cell Value", "Get Complete Row Info", "Get Rows For Range",
                "Execute Formula", "Add Formula To Cell", "Summation", "Vlookup",
                "Get Min Max Value From Row Data By Column Range",
                "Get Column Values For Keys In Sheet", "Add Rows With The List Of Entities",
                "Upload And Convert To PDF", "Convert To PDF", "Create Sheet",
                "Add Rows Async", "Get Result"})
@AuthScope(value = "Sites.ReadWrite.All", description = "Read and write all sites",
        catalogRequests = {"Add Table Row", "Get Table", "Get Table Rows", "Add Table Rows",
                "Delete Table Row", "Update Table Row", "Get Tables", "Get Drives",
                "Get Excel Files In Drive", "Get Excel Files In Folder", "Get Shared Files"})
@AuthScope(value = "User.Read", description = "Read user profile",
        catalogRequests = {"Test Connection"})
// ── Module 3: Audit Dashboard ──
@AuditConfig(
        dashboardTitle = "Office 365 Excel Audit Dashboard",
        enableCredentialHealth = true,
        enableCharts = true,
        csvExportPrefix = "excel-audit"
)
public class Office365ExcelAuthConfig {
}
