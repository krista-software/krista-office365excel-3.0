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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sheetoperations.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import static org.junit.jupiter.api.Assertions.assertSame;
import static org.mockito.Mockito.*;

/**
 * 100% branch coverage for the two {@code getGetRowsForRangeExtensionResponse} overloads
 * in {@link SheetOperationsFacade}.
 *
 * <p>Both overloads are pure delegation methods — no conditional logic.
 * The tests verify:
 * <ul>
 *   <li>Every argument is forwarded verbatim to the underlying {@code RowResponseService}</li>
 *   <li>The response returned by {@code RowResponseService} is passed back unchanged</li>
 *   <li>Non-null and null {@code driveId} are both forwarded correctly</li>
 *   <li>{@code allowRetry} = null / true / false are all forwarded correctly</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@DisplayName("SheetOperationsFacade – getGetRowsForRangeExtensionResponse – 100% branch coverage")
class SheetOperationsFacadeGetRowsForRangeTest {

    // ─── Test constants ──────────────────────────────────────────────────────
    private static final String FILE_ID    = "file-123";
    private static final String DRIVE_ID   = "drive-456";
    private static final String SHEET_NAME = "Sheet1";
    private static final String RANGE      = "A1:D10";

    // ─── Mocks ───────────────────────────────────────────────────────────────
    @Mock private CellResponseService      cellResponseService;
    @Mock private WorksheetResponseService worksheetResponseService;
    @Mock private RowResponseService        rowResponseService;
    @Mock private ColumnResponseService    columnResponseService;
    @Mock private FormulaResponseService   formulaResponseService;
    @Mock private SearchResponseService    searchResponseService;
    @Mock private WorkbookResponseService  workbookResponseService;
    @Mock private PdfResponseService       pdfResponseService;

    private SheetOperationsFacade facade;

    @BeforeEach
    void setUp() {
        facade = new SheetOperationsFacade(
                cellResponseService, worksheetResponseService, rowResponseService,
                columnResponseService, formulaResponseService, searchResponseService,
                workbookResponseService, pdfResponseService);
    }

    // ─── Helper ──────────────────────────────────────────────────────────────

    private ExtensionResponse mockResponse() {
        return mock(ExtensionResponse.class);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 4-arg overload
    // getGetRowsForRangeExtensionResponse(fileID, sheetName, range, allowRetry)
    //   → rowResponseService.getGetRowsForRangeExtensionResponse(fileID, sheetName, range, allowRetry)
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("4-arg overload — delegates to rowResponseService 4-arg")
    class FourArgOverload {

        @Test
        @DisplayName("allowRetry=null — all parameters forwarded, response returned unchanged")
        void allowRetryNull_forwardsAllParams_returnsServiceResponse() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, null))
                    .thenReturn(expected);

            ExtensionResponse result =
                    facade.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, null);

            assertSame(expected, result);
            verify(rowResponseService).getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, null);
        }

        @Test
        @DisplayName("allowRetry=true — forwarded to rowResponseService")
        void allowRetryTrue_forwarded() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, true))
                    .thenReturn(expected);

            ExtensionResponse result =
                    facade.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, true);

            assertSame(expected, result);
            verify(rowResponseService).getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, true);
        }

        @Test
        @DisplayName("allowRetry=false — forwarded to rowResponseService")
        void allowRetryFalse_forwarded() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, false))
                    .thenReturn(expected);

            ExtensionResponse result =
                    facade.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, false);

            assertSame(expected, result);
            verify(rowResponseService).getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, false);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 5-arg overload
    // getGetRowsForRangeExtensionResponse(driveId, fileID, sheetName, range, allowRetry)
    //   → rowResponseService.getGetRowsForRangeExtensionResponse(driveId, fileID, sheetName, range, allowRetry)
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("5-arg overload — delegates to rowResponseService 5-arg")
    class FiveArgOverload {

        @Test
        @DisplayName("non-null driveId, allowRetry=null — all parameters forwarded, response returned unchanged")
        void nonNullDriveId_allowRetryNull_forwardsAllParams() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null))
                    .thenReturn(expected);

            ExtensionResponse result =
                    facade.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

            assertSame(expected, result);
            verify(rowResponseService).getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);
        }

        @Test
        @DisplayName("null driveId — forwarded as null to rowResponseService")
        void nullDriveId_forwardedAsNull() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(null, FILE_ID, SHEET_NAME, RANGE, null))
                    .thenReturn(expected);

            ExtensionResponse result =
                    facade.getGetRowsForRangeExtensionResponse((String) null, FILE_ID, SHEET_NAME, RANGE, null);

            assertSame(expected, result);
            verify(rowResponseService).getGetRowsForRangeExtensionResponse(
                    null, FILE_ID, SHEET_NAME, RANGE, null);
        }

        @Test
        @DisplayName("allowRetry=true — forwarded to rowResponseService")
        void allowRetryTrue_forwarded() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true))
                    .thenReturn(expected);

            ExtensionResponse result =
                    facade.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

            assertSame(expected, result);
            verify(rowResponseService).getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);
        }

        @Test
        @DisplayName("allowRetry=false — forwarded to rowResponseService")
        void allowRetryFalse_forwarded() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, false))
                    .thenReturn(expected);

            ExtensionResponse result =
                    facade.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, false);

            assertSame(expected, result);
            verify(rowResponseService).getGetRowsForRangeExtensionResponse(
                    DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, false);
        }

        @Test
        @DisplayName("no other service method is called during delegation")
        void onlyRowResponseService_isCalled() {
            ExtensionResponse expected = mockResponse();
            when(rowResponseService.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null))
                    .thenReturn(expected);

            facade.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

            verifyNoInteractions(cellResponseService, worksheetResponseService, columnResponseService,
                    formulaResponseService, searchResponseService, workbookResponseService, pdfResponseService);
        }
    }
}
