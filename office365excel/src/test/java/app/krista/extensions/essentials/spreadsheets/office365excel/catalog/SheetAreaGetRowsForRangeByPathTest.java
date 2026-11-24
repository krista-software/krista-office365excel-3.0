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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.ExtensionResponse;
import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.SheetOperationsFacade;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.util.EventHandler;
import app.krista.ksdk.accounts.Account;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.ksdk.entities.Entities;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * 100% branch coverage for {@link SheetArea#getRowsForRangeByPath} and the
 * {@code audited()} private wrapper it delegates to.
 *
 * <p>Token validation logic (SheetArea pattern):
 * <pre>
 *   boolean useEmail = !requestContext.invokeAsUser();
 *   if (!useEmail) { providerFactory.create().validateToken(accountId); }
 * </pre>
 * i.e. invokeAsUser=true  → validateToken IS called
 *      invokeAsUser=false → validateToken is skipped
 *
 * <p>Branches covered:
 * <ul>
 *   <li>invokeAsUser=true  → providerFactory.create() called, validateToken called</li>
 *   <li>invokeAsUser=false → providerFactory.create() never called</li>
 *   <li>audited() result: SUCCESS response → CATALOG_REQUEST_SUCCESS audit</li>
 *   <li>audited() result: FAILURE response → CATALOG_REQUEST_FAILURE audit</li>
 *   <li>audited() exception: MustAuthorizeException → audit failure, re-throw</li>
 *   <li>audited() exception: RuntimeException → audit failure, re-throw as-is</li>
 *   <li>audited() exception: checked Exception → audit failure, wrapped in RuntimeException</li>
 *   <li>Parameter propagation: driveId, fileId, sheetName, range forwarded correctly</li>
 *   <li>Parameter propagation: allowRetry null, true, false all forwarded</li>
 *   <li>Parameter propagation: fileSearch forwarded to driveService</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("SheetArea – getRowsForRangeByPath – 100% branch coverage")
class SheetAreaGetRowsForRangeByPathTest {

    // ─── Test constants ──────────────────────────────────────────────────────
    private static final String SOURCE_PATH = "/Engineering/Shared Documents/Q1.xlsx";
    private static final String SHEET_NAME  = "Sheet1";
    private static final String RANGE       = "A1:D10";
    private static final String FILE_SEARCH = "My Drive";
    private static final String FILE_ID     = "file-123";
    private static final String DRIVE_ID    = "drive-456";
    private static final String ACCOUNT_ID  = "account-789";
    private static final String INVOKER_ID  = "test-invoker";

    // ─── Mocks ───────────────────────────────────────────────────────────────
    @Mock private Entities entities;
    @Mock private WorkbookService workbookService;
    @Mock private WorksheetService worksheetService;
    @Mock private RestClientProvider restClientProvider;
    @Mock private WorksheetOperation worksheetOperation;
    @Mock private StringUtil stringUtil;
    @Mock private EventHandler eventHandler;
    @Mock private SheetFormattingService sheetFormattingService;
    @Mock private SheetOperationsFacade sheetOperationsFacade;
    @Mock private DriveService driveService;
    @Mock private Office365ExcelAuditService auditService;
    @Mock private RequestContext requestContext;
    @Mock private AuthorizationContext authorizationContext;
    @Mock private GraphServiceClientProviderFactory providerFactory;
    @Mock private GraphServiceClientProvider clientProvider;
    @Mock private Invoker invoker;

    private SheetArea sheetArea;

    @BeforeEach
    void setUp() {
        when(invoker.getInvokerId()).thenReturn(INVOKER_ID);
        sheetArea = new SheetArea(
                entities, workbookService, worksheetService, restClientProvider,
                worksheetOperation, stringUtil, eventHandler, sheetFormattingService,
                sheetOperationsFacade, driveService, auditService,
                requestContext, authorizationContext, providerFactory, invoker);
    }

    // ─── Helpers ─────────────────────────────────────────────────────────────

    private ExcelFile excelFile() {
        return new ExcelFile(FILE_ID, DRIVE_ID, "Q1.xlsx", "https://example.com/Q1.xlsx");
    }

    /** Mock ExtensionResponse whose getResult() returns SUCCESS (not FAILURE → goes to success audit path). */
    private ExtensionResponse successResponse() {
        ExtensionResponse resp = mock(ExtensionResponse.class);
        when(resp.getResult()).thenReturn(ExtensionResponse.Result.SUCCESS);
        return resp;
    }

    /** Mock ExtensionResponse whose getResult() returns FAILURE. */
    private ExtensionResponse failureResponse() {
        ExtensionResponse resp = mock(ExtensionResponse.class);
        when(resp.getResult()).thenReturn(ExtensionResponse.Result.FAILURE);
        return resp;
    }

    /**
     * Stubs authorizationContext + providerFactory so that the
     * invokeAsUser=true path (which calls validateToken) does not NPE.
     */
    private void stubInvokeAsUserTrue() {
        when(requestContext.invokeAsUser()).thenReturn(true);
        Account account = mock(Account.class);
        when(authorizationContext.getAuthorizedAccount()).thenReturn(account);
        when(account.getAccountId()).thenReturn(ACCOUNT_ID);
        when(providerFactory.create()).thenReturn(clientProvider);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Token Validation Branch (invokeAsUser)
    //
    // SheetArea pattern:
    //   boolean useEmail = !requestContext.invokeAsUser();
    //   if (!useEmail) { providerFactory.create().validateToken(accountId); }
    //
    // invokeAsUser=true  → useEmail=false → !useEmail=true  → validateToken IS called
    // invokeAsUser=false → useEmail=true  → !useEmail=false → validateToken is skipped
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Token validation (invokeAsUser branch)")
    class TokenValidation {

        @Test
        @DisplayName("invokeAsUser=true → providerFactory.create() called, validateToken called with accountId")
        void invokeAsUser_true_callsValidateToken() {
            stubInvokeAsUserTrue();
            when(driveService.getExcelFileByPath(SOURCE_PATH, FILE_SEARCH)).thenReturn(excelFile());
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null)).thenReturn(resp);

            ExtensionResponse result = sheetArea.getRowsForRangeByPath(
                    SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null);

            assertSame(resp, result);
            verify(providerFactory).create();
            verify(clientProvider).validateToken(ACCOUNT_ID);
        }

        @Test
        @DisplayName("invokeAsUser=false → providerFactory.create() is never called")
        void invokeAsUser_false_skipsTokenValidation() {
            when(requestContext.invokeAsUser()).thenReturn(false);
            when(driveService.getExcelFileByPath(SOURCE_PATH, FILE_SEARCH)).thenReturn(excelFile());
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null)).thenReturn(resp);

            sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null);

            verify(providerFactory, never()).create();
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // audited() — Result Audit Branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("audited() — result audit branches")
    class AuditedResult {

        @BeforeEach
        void stubCommon() {
            // invokeAsUser=false skips token validation — keeps these tests focused on audit logic
            when(requestContext.invokeAsUser()).thenReturn(false);
            when(driveService.getExcelFileByPath(SOURCE_PATH, FILE_SEARCH)).thenReturn(excelFile());
        }

        @Test
        @DisplayName("SUCCESS response → auditService.startRequest called, then CATALOG_REQUEST_SUCCESS recorded")
        void successResponse_auditsSuccess() {
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null)).thenReturn(resp);

            ExtensionResponse result = sheetArea.getRowsForRangeByPath(
                    SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null);

            assertSame(resp, result);
            verify(auditService).startRequest(INVOKER_ID, "Get Rows For Range By Path");
            verify(auditService).record(INVOKER_ID, "CATALOG_REQUEST_SUCCESS", "success",
                    "Get Rows For Range By Path", null, null, null);
        }

        @Test
        @DisplayName("FAILURE response → auditService records CATALOG_REQUEST_FAILURE (not success)")
        void failureResponse_auditsFailure() {
            ExtensionResponse resp = failureResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null)).thenReturn(resp);

            ExtensionResponse result = sheetArea.getRowsForRangeByPath(
                    SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null);

            assertSame(resp, result);
            verify(auditService).startRequest(INVOKER_ID, "Get Rows For Range By Path");
            verify(auditService).record(INVOKER_ID, "CATALOG_REQUEST_FAILURE", "failure",
                    "Get Rows For Range By Path", null, null, null);
            verify(auditService, never()).record(eq(INVOKER_ID), eq("CATALOG_REQUEST_SUCCESS"),
                    any(), any(), any(), any(), any());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // audited() — Exception Branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("audited() — exception branches")
    class AuditedExceptions {

        @BeforeEach
        void stubRequestContext() {
            // invokeAsUser=false skips token validation — keeps these tests focused on exception handling
            when(requestContext.invokeAsUser()).thenReturn(false);
        }

        @Test
        @DisplayName("MustAuthorizeException → audit failure recorded, exception re-thrown unchanged")
        void mustAuthorizeException_recordsFailureAndRethrows() {
            MustAuthorizeException ex = mock(MustAuthorizeException.class);
            when(ex.getMessage()).thenReturn("must authorize");
            doThrow(ex).when(driveService).getExcelFileByPath(any(), any());

            MustAuthorizeException thrown = assertThrows(MustAuthorizeException.class,
                    () -> sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null));

            assertSame(ex, thrown);
            verify(auditService).record(eq(INVOKER_ID), eq("CATALOG_REQUEST_FAILURE"), eq("failure"),
                    eq("Get Rows For Range By Path"), eq("must authorize"), eq("MustAuthorizeException"), isNull());
        }

        @Test
        @DisplayName("RuntimeException → audit failure recorded, same exception re-thrown")
        void runtimeException_recordsFailureAndRethrowsSameInstance() {
            IllegalArgumentException ex = new IllegalArgumentException("file not found");
            doThrow(ex).when(driveService).getExcelFileByPath(any(), any());

            RuntimeException thrown = assertThrows(RuntimeException.class,
                    () -> sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null));

            assertSame(ex, thrown);
            verify(auditService).record(eq(INVOKER_ID), eq("CATALOG_REQUEST_FAILURE"), eq("failure"),
                    eq("Get Rows For Range By Path"),
                    eq("file not found"), eq("IllegalArgumentException"), isNull());
        }

        @Test
        @DisplayName("Checked Exception (non-RuntimeException) → audit failure recorded, wrapped in RuntimeException")
        void checkedException_recordsFailureAndWrapsInRuntimeException() {
            Exception checked = new Exception("io failure");
            // doAnswer can throw any exception regardless of the method's declared throws clause
            doAnswer(inv -> { throw checked; }).when(driveService).getExcelFileByPath(any(), any());

            RuntimeException thrown = assertThrows(RuntimeException.class,
                    () -> sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null));

            assertSame(checked, thrown.getCause());
            verify(auditService).record(eq(INVOKER_ID), eq("CATALOG_REQUEST_FAILURE"), eq("failure"),
                    eq("Get Rows For Range By Path"),
                    eq("io failure"), eq("Exception"), isNull());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Parameter Propagation
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Parameter propagation")
    class ParameterPropagation {

        @BeforeEach
        void stubRequestContext() {
            // invokeAsUser=false skips token validation — keeps these tests focused on parameter routing
            when(requestContext.invokeAsUser()).thenReturn(false);
        }

        @Test
        @DisplayName("driveId and fileId resolved from ExcelFile are forwarded to facade")
        void driveIdAndFileId_fromExcelFile_forwardedToFacade() {
            when(driveService.getExcelFileByPath(SOURCE_PATH, FILE_SEARCH)).thenReturn(excelFile());
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null)).thenReturn(resp);

            sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, null);

            verify(sheetOperationsFacade).getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);
        }

        @Test
        @DisplayName("allowRetry=true forwarded to facade")
        void allowRetry_true_forwarded() {
            when(driveService.getExcelFileByPath(SOURCE_PATH, FILE_SEARCH)).thenReturn(excelFile());
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true)).thenReturn(resp);

            sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, true);

            verify(sheetOperationsFacade).getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);
        }

        @Test
        @DisplayName("allowRetry=false forwarded to facade")
        void allowRetry_false_forwarded() {
            when(driveService.getExcelFileByPath(SOURCE_PATH, FILE_SEARCH)).thenReturn(excelFile());
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, false)).thenReturn(resp);

            sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, FILE_SEARCH, false);

            verify(sheetOperationsFacade).getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, false);
        }

        @Test
        @DisplayName("fileSearch value forwarded verbatim to driveService.getExcelFileByPath")
        void fileSearch_forwarded_todriveService() {
            String siteFilesSearch = "Site Files";
            when(driveService.getExcelFileByPath(SOURCE_PATH, siteFilesSearch)).thenReturn(excelFile());
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null)).thenReturn(resp);

            sheetArea.getRowsForRangeByPath(SOURCE_PATH, SHEET_NAME, RANGE, siteFilesSearch, null);

            verify(driveService).getExcelFileByPath(SOURCE_PATH, siteFilesSearch);
        }

        @Test
        @DisplayName("sourcePath forwarded verbatim to driveService.getExcelFileByPath")
        void sourcePath_forwarded_todriveService() {
            String customPath = "/Finance/Budget 2026.xlsx";
            when(driveService.getExcelFileByPath(customPath, FILE_SEARCH)).thenReturn(excelFile());
            ExtensionResponse resp = successResponse();
            when(sheetOperationsFacade.getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null)).thenReturn(resp);

            sheetArea.getRowsForRangeByPath(customPath, SHEET_NAME, RANGE, FILE_SEARCH, null);

            verify(driveService).getExcelFileByPath(customPath, FILE_SEARCH);
        }
    }
}
