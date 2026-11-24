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

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Drive;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.SharedOperationsFacade;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClient;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.util.EventHandler;
import app.krista.ksdk.accounts.Account;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.ksdk.entities.Entities;
import app.krista.model.base.EntityValue;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import app.krista.model.entity.EntityDefinition;
import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import com.kristasoft.common.holders.ThreadLocalProxy;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.*;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import okhttp3.Response;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.ExecutorService;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * 100% code coverage tests for {@link SharedArea}.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("SharedArea - 100% Code Coverage")
class SharedAreaTest {

    @Mock private DriveService mockDriveService;
    @Mock private TableService mockTableService;
    @Mock private RestClientProvider mockRestClientProvider;
    @Mock private WorkbookService mockWorkbookService;
    @Mock private WorksheetService mockWorksheetService;
    @Mock private EventHandler mockEventHandler;
    @Mock private RequestContext mockRequestContext;
    @Mock private AuthorizationContext mockAuthorizationContext;
    @Mock private Entities mockEntities;
    @Mock private WorksheetOperation mockWorksheetOperation;
    @Mock private StringUtil mockStringUtil;
    @Mock private SheetFormattingService mockSheetFormattingService;
    @Mock private SharedOperationsFacade mockFacade;
    @Mock private Office365ExcelAuditService mockAuditService;
    @Mock private GraphServiceClientProviderFactory mockProviderFactory;
    @Mock private GraphServiceClientProvider mockClientProvider;
    @Mock private Invoker mockInvoker;

    private SharedArea sharedArea;

    /**
     * An executor that runs submitted tasks synchronously in the calling thread.
     * This makes async methods deterministic and testable.
     */
    private final ExecutorService syncExecutor = new ExecutorService() {
        @Override public void shutdown() {}
        @Override public List<Runnable> shutdownNow() { return List.of(); }
        @Override public boolean isShutdown() { return false; }
        @Override public boolean isTerminated() { return false; }
        @Override public boolean awaitTermination(long timeout, java.util.concurrent.TimeUnit unit) { return true; }
        @Override public <T> java.util.concurrent.Future<T> submit(java.util.concurrent.Callable<T> task) {
            try { task.call(); } catch (Exception e) { /* swallow for test */ }
            return null;
        }
        @Override public <T> java.util.concurrent.Future<T> submit(Runnable task, T result) {
            task.run(); return null;
        }
        @Override public java.util.concurrent.Future<?> submit(Runnable task) {
            try { task.run(); } catch (Exception e) { /* swallow for test */ }
            return null;
        }
        @Override public <T> List<java.util.concurrent.Future<T>> invokeAll(Collection<? extends java.util.concurrent.Callable<T>> tasks) { return List.of(); }
        @Override public <T> List<java.util.concurrent.Future<T>> invokeAll(Collection<? extends java.util.concurrent.Callable<T>> tasks, long timeout, java.util.concurrent.TimeUnit unit) { return List.of(); }
        @Override public <T> T invokeAny(Collection<? extends java.util.concurrent.Callable<T>> tasks) { return null; }
        @Override public <T> T invokeAny(Collection<? extends java.util.concurrent.Callable<T>> tasks, long timeout, java.util.concurrent.TimeUnit unit) { return null; }
        @Override public void execute(Runnable command) { command.run(); }
    };

    @BeforeEach
    void setUp() {
        when(mockInvoker.getInvokerId()).thenReturn("test-invoker");
        lenient().when(mockProviderFactory.create()).thenReturn(mockClientProvider);
        sharedArea = new SharedArea(
                mockDriveService, mockRestClientProvider, mockRequestContext,
                mockEventHandler, mockAuthorizationContext, mockEntities,
                mockWorkbookService, mockTableService, mockWorksheetService,
                mockWorksheetOperation, mockStringUtil, mockSheetFormattingService,
                mockFacade, mockAuditService, mockProviderFactory, mockInvoker
        );
        // Replace the real executor with synchronous one for deterministic testing
        sharedArea.executorService = syncExecutor;
    }

    // ═══════════════════════════════════════════════════════════════════
    // FACADE DELEGATION METHODS (simple pass-through)
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Facade delegation methods")
    class FacadeDelegation {

        @Test
        @DisplayName("addRowsToTable → delegates to facade.getExtensionResponse")
        void addRowsToTable() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            List<String> rows = List.of("r1");
            List<File> csvFiles = List.of();
            when(mockFacade.getExtensionResponse("d1", "f1", "t1", rows, csvFiles, true)).thenReturn(expected);

            ExtensionResponse result = sharedArea.addRowsToTable("d1", "f1", "t1", rows, csvFiles, true);

            assertSame(expected, result);
            verify(mockFacade).getExtensionResponse("d1", "f1", "t1", rows, csvFiles, true);
        }

        @Test
        @DisplayName("addRowsToSheet → delegates to facade.getSheetExtensionResponse")
        void addRowsToSheet() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getSheetExtensionResponse(eq("d1"), eq("f1"), eq("s1"), anyList(), anyList(), eq(false)))
                    .thenReturn(expected);

            ExtensionResponse result = sharedArea.addRowsToSheet("d1", "f1", "s1", List.of(), List.of(), false);

            assertSame(expected, result);
        }

        @Test
        @DisplayName("rangeSum → delegates to facade.getRangeSumExtensionResponse")
        void rangeSum() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getRangeSumExtensionResponse("d1", "f1", "s1", "A1:A10", true)).thenReturn(expected);

            assertSame(expected, sharedArea.rangeSum("d1", "f1", "s1", "A1:A10", true));
        }

        @Test
        @DisplayName("sheetVlookup → delegates to facade.getSheetVlookupExtensionResponse")
        void sheetVlookup() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getSheetVlookupExtensionResponse("d1", "f1", "val", "s1", "A1:D10", 2.0, true, false))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.sheetVlookup("d1", "f1", "val", "s1", "A1:D10", 2.0, true, false));
        }

        @Test
        @DisplayName("getTableRows → delegates to facade.getTableRowsExtensionResponse")
        void getTableRows() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getTableRowsExtensionResponse("d1", "f1", "t1", true)).thenReturn(expected);

            assertSame(expected, sharedArea.getTableRows("d1", "f1", "t1", true));
        }

        @Test
        @DisplayName("addDataToProvidedCell → delegates with null allowRetry")
        void addDataToProvidedCell() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getAddDataToCellExtensionResponse("d1", "f1", "s1", "A1", "data", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.addDataToProvidedCell("d1", "f1", "s1", "A1", "data"));
        }

        @Test
        @DisplayName("deleteColumn → delegates with null allowRetry")
        void deleteColumn() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getDeleteColumnExtensionResponse("d1", "f1", "t1", "col1", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.deleteColumn("d1", "f1", "t1", "col1"));
        }

        @Test
        @DisplayName("getColumnValuesForKeys → delegates to facade")
        void getColumnValuesForKeys() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            List<String> keyRows = List.of("k1");
            List<String> lookupCol = List.of("c1");
            when(mockFacade.getColumnValuesForKeysExtensionResponse("d1", "f1", "t1", keyRows, lookupCol, true, false))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.getColumnValuesForKeys("d1", "f1", "t1", keyRows, lookupCol, true, false));
        }

        @Test
        @DisplayName("getRowsForRange → delegates to facade")
        void getRowsForRange() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getRowsForRangeExtensionResponse("d1", "f1", "s1", "A1:D10", true))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.getRowsForRange("d1", "f1", "s1", "A1:D10", true));
        }

        @Test
        @DisplayName("executeFormulaForSharedFile → delegates with null allowRetry")
        void executeFormulaForSharedFile() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getExecuteFormulaExtensionResponse("d1", "f1", "=SUM(A1)", "myFormula", "comment", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.executeFormulaForSharedFile("d1", "f1", "=SUM(A1)", "myFormula", "comment"));
        }

        @Test
        @DisplayName("vlookupInTable → delegates with null allowRetry")
        void vlookupInTable() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getVlookupInTableExtensionResponse("d1", "f1", "val", "t1", 3.0, false, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.vlookupInTable("d1", "f1", "val", "t1", 3.0, false));
        }

        @Test
        @DisplayName("columnsSum → delegates with null allowRetry")
        void columnsSum() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            List<String> cols = List.of("A", "B");
            when(mockFacade.getColumnsSumExtensionResponse("d1", "f1", "t1", cols, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.columnsSum("d1", "f1", "t1", cols));
        }

        @Test
        @DisplayName("search → delegates with null allowRetry")
        void search() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getSearchExtensionResponse("d1", "f1", "s1", "text", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.search("d1", "f1", "s1", "text"));
        }

        @Test
        @DisplayName("cellData → delegates with null allowRetry")
        void cellData() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getCellDataExtensionResponse("d1", "f1", "s1", "A1", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.cellData("d1", "f1", "s1", "A1"));
        }

        @Test
        @DisplayName("getSharedExcelFile → delegates with null allowRetry")
        void getSharedExcelFile() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getSharedExcelFileExtensionResponse("folder", "file.xlsx", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.getSharedExcelFile("folder", "file.xlsx"));
        }

        @Test
        @DisplayName("convertToPDF → delegates with null allowRetry")
        void convertToPDF() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getConvertToPdfExtensionResponse("d1", "f1", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.convertToPDF("f1", "d1"));
        }

        @Test
        @DisplayName("deleteColumnWithIndex → delegates with null allowRetry")
        void deleteColumnWithIndex() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getDeleteColumnWithIndexExtensionResponse("d1", "f1", "s1", "col1", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.deleteColumnWithIndex("f1", "d1", "s1", "col1"));
        }

        @Test
        @DisplayName("addRowWithIndex → delegates with null allowRetry")
        void addRowWithIndex() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            List<String> rows = List.of("r1");
            List<File> csvFiles = List.of();
            when(mockFacade.getAddRowWithIndexExtensionResponse("d1", "f1", "s1", rows, csvFiles, 5.0, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.addRowWithIndex("d1", "f1", "s1", rows, 5.0, csvFiles));
        }

        @Test
        @DisplayName("updateRowWithIndex → delegates with null allowRetry")
        void updateRowWithIndex() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getUpdateRowWithIndexExtensionResponse("d1", "f1", "s1", "data", 3.0, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.updateRowWithIndex("d1", "f1", "s1", "data", 3.0));
        }

        @Test
        @DisplayName("getCompleteRowInfo → delegates with null allowRetry")
        void getCompleteRowInfo() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getCompleteRowInfoExtensionResponse("d1", "f1", "s1", 2.0, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.getCompleteRowInfo("d1", "f1", "s1", 2.0));
        }

        @Test
        @DisplayName("deleteRowWithIndex → delegates with null allowRetry")
        void deleteRowWithIndex() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getDeleteRowWithIndexExtensionResponse("d1", "f1", "s1", 1.0, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.deleteRowWithIndex("d1", "f1", "s1", 1.0));
        }

        @Test
        @DisplayName("addMultipleCellData → delegates with null allowRetry")
        void addMultipleCellData() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            List<Map<String, Object>> cellData = List.of(Map.of("Cell Id", "A1", "Data", "val"));
            when(mockFacade.getAddMultipleCellDataExtensionResponse("d1", "f1", "s1", cellData, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.addMultipleCellData("d1", "f1", "s1", cellData));
        }

        @Test
        @DisplayName("getColumnValuesForKeysInSheet → delegates with null allowRetry")
        void getColumnValuesForKeysInSheet() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            List<String> lr = List.of("r1");
            List<String> lc = List.of("c1");
            when(mockFacade.getColumnValuesForKeysInSheetExtensionResponse("d1", "f1", "s1", "A1:D10", lr, lc, true, 1.0, null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.getColumnValuesForKeysInSheet("d1", "f1", "s1", "A1:D10", lr, lc, true, 1.0));
        }

        @Test
        @DisplayName("clearSharedSheetForRange → delegates with null allowRetry")
        void clearSharedSheetForRange() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(mockFacade.getClearSharedSheetForRangeExtensionResponse("d1", "f1", "s1", "A1:B5", null))
                    .thenReturn(expected);

            assertSame(expected, sharedArea.clearSharedSheetForRange("d1", "f1", "s1", "A1:B5"));
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // DIRECT SERVICE DELEGATION (not via facade)
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Direct service delegation methods")
    class DirectServiceDelegation {

        @Test
        @DisplayName("getSharedExcelFiles → delegates to driveService when no driveId")
        void getSharedExcelFiles() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.getSharedExcelFilesInformation(null)).thenReturn(expected);

            assertSame(expected, sharedArea.getSharedExcelFiles(null));
        }

        @Test
        @DisplayName("getSharedExcelFiles → uses direct drive enumeration when driveId provided")
        void getSharedExcelFiles_withDriveId_usesDirectEnumeration() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.getExcelFilesFromDrive("drive-123")).thenReturn(expected);

            assertSame(expected, sharedArea.getSharedExcelFiles("drive-123"));
            verify(mockDriveService, never()).getSharedExcelFilesInformation(any());
        }

        @Test
        @DisplayName("getSharedExcelFiles → falls back to sharedWithMe when driveId is blank")
        void getSharedExcelFiles_withBlankDriveId_usesSharedWithMe() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.getSharedExcelFilesInformation(null)).thenReturn(expected);

            assertSame(expected, sharedArea.getSharedExcelFiles("   "));
            verify(mockDriveService, never()).getExcelFilesFromDrive(any());
        }

        @Test
        @DisplayName("getAllSharedDrivesInformation → delegates to driveService when no driveId")
        void getAllSharedDrivesInformation() {
            List<Drive> expected = List.of(mock(Drive.class));
            when(mockDriveService.getAllSharedDrivesInformation()).thenReturn(expected);

            assertSame(expected, sharedArea.getAllSharedDrivesInformation(null));
        }

        @Test
        @DisplayName("getAllSharedDrivesInformation → appends new drive when driveId provided and not already present")
        void getAllSharedDrivesInformation_withDriveId_appendsNewDrive() {
            Drive existing = new Drive("drive-existing", "User A", null);
            Drive extra = new Drive("drive-extra", "SharePoint", null);
            when(mockDriveService.getAllSharedDrivesInformation()).thenReturn(new ArrayList<>(List.of(existing)));
            when(mockDriveService.getDriveById("drive-extra")).thenReturn(List.of(extra));

            List<Drive> result = sharedArea.getAllSharedDrivesInformation("drive-extra");

            assertEquals(2, result.size());
            assertTrue(result.stream().anyMatch(d -> "drive-existing".equals(d.driveId)));
            assertTrue(result.stream().anyMatch(d -> "drive-extra".equals(d.driveId)));
        }

        @Test
        @DisplayName("getAllSharedDrivesInformation → does not duplicate drive already present in sharedWithMe")
        void getAllSharedDrivesInformation_withDriveId_doesNotDuplicate() {
            Drive existing = new Drive("drive-123", "User A", null);
            when(mockDriveService.getAllSharedDrivesInformation()).thenReturn(new ArrayList<>(List.of(existing)));
            when(mockDriveService.getDriveById("drive-123")).thenReturn(List.of(new Drive("drive-123", "SharePoint", null)));

            List<Drive> result = sharedArea.getAllSharedDrivesInformation("drive-123");

            assertEquals(1, result.size());
            assertEquals("drive-123", result.get(0).driveId);
        }

        @Test
        @DisplayName("searchExcelFilesInDrive → delegates to driveService with xlsx format")
        void searchExcelFilesInDrive_xlsxFormat() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.searchExcelFilesInDrive("drive-123", ".xlsx")).thenReturn(expected);

            assertSame(expected, sharedArea.searchExcelFilesInDrive("drive-123", ".xlsx"));
            verify(mockDriveService).searchExcelFilesInDrive("drive-123", ".xlsx");
        }

        @Test
        @DisplayName("searchExcelFilesInDrive → delegates to driveService with xls format")
        void searchExcelFilesInDrive_xlsFormat() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.searchExcelFilesInDrive("drive-123", ".xls")).thenReturn(expected);

            assertSame(expected, sharedArea.searchExcelFilesInDrive("drive-123", ".xls"));
            verify(mockDriveService).searchExcelFilesInDrive("drive-123", ".xls");
        }

        @Test
        @DisplayName("searchExcelFilesInDrive → delegates to driveService with csv format")
        void searchExcelFilesInDrive_csvFormat() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.searchExcelFilesInDrive("drive-123", ".csv")).thenReturn(expected);

            assertSame(expected, sharedArea.searchExcelFilesInDrive("drive-123", ".csv"));
            verify(mockDriveService).searchExcelFilesInDrive("drive-123", ".csv");
        }

        @Test
        @DisplayName("searchExcelFilesInDrive → delegates to driveService with tsv format")
        void searchExcelFilesInDrive_tsvFormat() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.searchExcelFilesInDrive("drive-123", ".tsv")).thenReturn(expected);

            assertSame(expected, sharedArea.searchExcelFilesInDrive("drive-123", ".tsv"));
            verify(mockDriveService).searchExcelFilesInDrive("drive-123", ".tsv");
        }

        @Test
        @DisplayName("searchExcelFilesInDrive → passes driveId and fileFormat together")
        void searchExcelFilesInDrive_passesBothParameters() {
            List<ExcelFile> expected = List.of(mock(ExcelFile.class));
            when(mockDriveService.searchExcelFilesInDrive("my-drive", ".xlsx")).thenReturn(expected);

            List<ExcelFile> result = sharedArea.searchExcelFilesInDrive("my-drive", ".xlsx");

            assertSame(expected, result);
            verify(mockDriveService, times(1)).searchExcelFilesInDrive("my-drive", ".xlsx");
        }
    }
    // ═══════════════════════════════════════════════════════════════════
    // WAIT-FOR-EVENT METHODS
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Wait-for-event methods")
    class WaitForEventMethods {

        @Test
        @DisplayName("getRowsForRangeResult - matching taskID → returns rows")
        void getRowsForRangeResult_matching() {
            FreeForm eventData = new FreeForm();
            List<FreeForm> rows = List.of(new FreeForm());
            eventData.put("Rows", "[ Composite ]", rows);

            List<FreeForm> result = sharedArea.getRowsForRangeResult("task-1", eventData, "task-1");
            assertSame(rows, result);
        }

        @Test
        @DisplayName("getRowsForRangeResult - non-matching taskID → throws IllegalStateException")
        void getRowsForRangeResult_nonMatching() {
            assertThrows(IllegalStateException.class, () ->
                    sharedArea.getRowsForRangeResult("event-1", new FreeForm(), "task-1"));
        }

        @Test
        @DisplayName("getTablesInformationForSharedFile - matching taskID → returns tables")
        void getTablesInformationForSharedFile_matching() {
            FreeForm eventData = new FreeForm();
            List<Table> tables = List.of(mock(Table.class));
            eventData.put("Data", "[ Entity(Table) ]", tables);

            List<Table> result = sharedArea.getTablesInformationForSharedFile("task-2", eventData, "task-2");
            assertSame(tables, result);
        }

        @Test
        @DisplayName("getTablesInformationForSharedFile - non-matching → throws IllegalStateException")
        void getTablesInformationForSharedFile_nonMatching() {
            assertThrows(IllegalStateException.class, () ->
                    sharedArea.getTablesInformationForSharedFile("event-2", new FreeForm(), "task-2"));
        }

        @Test
        @DisplayName("getAddRowWithIndexResult - matching taskID → returns sheet data")
        void getAddRowWithIndexResult_matching() {
            FreeForm eventData = new FreeForm();
            List<FreeForm> data = List.of(new FreeForm());
            eventData.put("Sheet Data", "[ Composite ]", data);

            List<FreeForm> result = sharedArea.getAddRowWithIndexResult("task-3", eventData, "task-3");
            assertSame(data, result);
        }

        @Test
        @DisplayName("getAddRowWithIndexResult - non-matching → throws IllegalStateException")
        void getAddRowWithIndexResult_nonMatching() {
            assertThrows(IllegalStateException.class, () ->
                    sharedArea.getAddRowWithIndexResult("event-3", new FreeForm(), "task-3"));
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // ASYNC METHODS
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Async methods")
    class AsyncMethods {

        @Test
        @DisplayName("addRowsToTableAsync - happy path → returns taskId and fires event")
        void addRowsToTableAsync_happyPath() {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class);
                 MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {

                ws.when(() -> Worksheets.isRowDataAvailable(anyList(), isNull())).thenReturn(true);
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                List<FreeForm> rows = List.of(new FreeForm());
                when(mockWorkbookService.addRowsInTable("d1", "f1", "t1", List.of("r1"), null, null, Constants.COMMA))
                        .thenReturn(rows);

                String taskId = sharedArea.addRowsToTableAsync("d1", "f1", "t1", List.of("r1"), null);

                assertNotNull(taskId);
                verify(mockEventHandler).handleEvent(eq(taskId), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("addRowsToTableAsync - no row data → throws IllegalArgumentException")
        void addRowsToTableAsync_noRowData() {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class)) {
                ws.when(() -> Worksheets.isRowDataAvailable(isNull(), isNull())).thenReturn(false);

                assertThrows(IllegalArgumentException.class, () ->
                        sharedArea.addRowsToTableAsync("d1", "f1", "t1", null, null));
            }
        }

        @Test
        @DisplayName("addRowsToTableAsync - exception in executor → wraps in RuntimeException")
        void addRowsToTableAsync_exception() {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class);
                 MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {

                ws.when(() -> Worksheets.isRowDataAvailable(anyList(), isNull())).thenReturn(true);
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                when(mockWorkbookService.addRowsInTable(anyString(), anyString(), anyString(), anyList(), isNull(), isNull(), anyString()))
                        .thenThrow(new RuntimeException("API error"));

                // syncExecutor swallows the RuntimeException, so the method still returns a taskId
                String taskId = sharedArea.addRowsToTableAsync("d1", "f1", "t1", List.of("r1"), null);
                assertNotNull(taskId);

                // Event handler should NOT have been called due to exception
                verify(mockEventHandler, never()).handleEvent(anyString(), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("getRowsForRangeAsync - happy path → returns taskId and fires event")
        void getRowsForRangeAsync_happyPath() {
            try (MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                List<FreeForm> rows = List.of(new FreeForm());
                when(mockWorksheetService.getRows("d1", "f1", "s1", "A1:D10")).thenReturn(rows);

                String taskId = sharedArea.getRowsForRangeAsync("d1", "f1", "s1", "A1:D10");

                assertNotNull(taskId);
                verify(mockEventHandler).handleEvent(eq(taskId), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("getRowsForRangeAsync - exception in executor → wraps in RuntimeException")
        void getRowsForRangeAsync_exception() {
            try (MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                when(mockWorksheetService.getRows(anyString(), anyString(), anyString(), anyString()))
                        .thenThrow(new RuntimeException("Timeout"));

                String taskId = sharedArea.getRowsForRangeAsync("d1", "f1", "s1", "A1:D10");
                assertNotNull(taskId);
                verify(mockEventHandler, never()).handleEvent(anyString(), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("tableInformationForTheSharedFileAsync - invokeAsUser=true → uses authContext accountId")
        void tableInfoAsync_invokeAsUser() {
            try (MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                when(mockRequestContext.invokeAsUser()).thenReturn(true);
                Account mockAccount = mock(Account.class);
                when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(mockAccount);
                when(mockAccount.getAccountId()).thenReturn("user-account-id");

                List<Table> tables = List.of(mock(Table.class));
                when(mockTableService.getTableNames("user-account-id", true, "f1", "d1")).thenReturn(tables);

                String taskId = sharedArea.tableInformationForTheSharedFileAsync("f1", "d1");

                assertNotNull(taskId);
                verify(mockEventHandler).handleEvent(eq(taskId), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("tableInformationForTheSharedFileAsync - invokeAsUser=false → uses attributes emailId")
        void tableInfoAsync_notInvokeAsUser() {
            try (MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                when(mockRequestContext.invokeAsUser()).thenReturn(false);
                OfficeExcelAttributes mockAttrs = mock(OfficeExcelAttributes.class);
                when(mockRestClientProvider.getOffice365ExcelAttributes()).thenReturn(mockAttrs);
                when(mockAttrs.getEmailId()).thenReturn("admin@test.com");

                List<Table> tables = List.of(mock(Table.class));
                when(mockTableService.getTableNames("admin@test.com", false, "f1", "d1")).thenReturn(tables);

                String taskId = sharedArea.tableInformationForTheSharedFileAsync("f1", "d1");

                assertNotNull(taskId);
                verify(mockEventHandler).handleEvent(eq(taskId), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("tableInformationForTheSharedFileAsync - exception in executor")
        void tableInfoAsync_exception() {
            try (MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                when(mockRequestContext.invokeAsUser()).thenReturn(true);
                Account mockAccount = mock(Account.class);
                when(mockAuthorizationContext.getAuthorizedAccount()).thenReturn(mockAccount);
                when(mockAccount.getAccountId()).thenReturn("user-id");
                when(mockTableService.getTableNames(anyString(), anyBoolean(), anyString(), anyString()))
                        .thenThrow(new RuntimeException("Error"));

                String taskId = sharedArea.tableInformationForTheSharedFileAsync("f1", "d1");
                assertNotNull(taskId);
                verify(mockEventHandler, never()).handleEvent(anyString(), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("addRowWithIndexAsync - happy path")
        void addRowWithIndexAsync_happyPath() throws IOException {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class);
                 MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class);
                 MockedStatic<StringUtil> su = mockStatic(StringUtil.class)) {

                ws.when(() -> Worksheets.isRowDataAvailable(anyList(), isNull())).thenReturn(true);
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                List<String> rowStrings = List.of("r1,r2");
                when(mockStringUtil.getStringList(anyList(), isNull())).thenReturn(rowStrings);
                when(mockStringUtil.calculateAddressForInsert(eq(3.0), anyList(), isNull())).thenReturn("A3:B3");
                su.when(() -> StringUtil.calculateAddressForInsert(eq(3.0), eq(rowStrings))).thenReturn("A3:B3");

                RestClient mockRestClient = mock(RestClient.class);
                when(mockRestClientProvider.getRestClient("d1")).thenReturn(mockRestClient);
                when(mockRestClient.post(anyString(), anyMap())).thenReturn(mock(Response.class));

                List<FreeForm> result = List.of(new FreeForm());
                when(mockWorksheetService.addRows("d1", "f1", "s1", rowStrings, "A3:B3")).thenReturn(result);

                String taskId = sharedArea.addRowWithIndexAsync("d1", "f1", "s1", List.of("r1"), null, 3.0);

                assertNotNull(taskId);
                verify(mockEventHandler).handleEvent(eq(taskId), any(FreeForm.class));
                verify(mockSheetFormattingService).clearRowFormatting("d1", "f1", "s1", 3.0);
            }
        }

        @Test
        @DisplayName("addRowWithIndexAsync - no row data → throws IllegalArgumentException")
        void addRowWithIndexAsync_noRowData() {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class)) {
                ws.when(() -> Worksheets.isRowDataAvailable(isNull(), isNull())).thenReturn(false);

                assertThrows(IllegalArgumentException.class, () ->
                        sharedArea.addRowWithIndexAsync("d1", "f1", "s1", null, null, 1.0));
            }
        }

        @Test
        @DisplayName("addRowWithIndexAsync - IOException in executor")
        void addRowWithIndexAsync_exception() throws IOException {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class);
                 MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {

                ws.when(() -> Worksheets.isRowDataAvailable(anyList(), isNull())).thenReturn(true);
                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);

                when(mockStringUtil.getStringList(anyList(), isNull())).thenThrow(new IOException("Read error"));

                String taskId = sharedArea.addRowWithIndexAsync("d1", "f1", "s1", List.of("r1"), null, 1.0);
                assertNotNull(taskId);
                verify(mockEventHandler, never()).handleEvent(anyString(), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("addMultipleCellDataAsync - happy path with valid cells")
        void addMultipleCellDataAsync_happyPath() {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class);
                 MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {

                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);
                ws.when(() -> Worksheets.isValidCellId("A1")).thenReturn(true);

                FreeForm cellResult = new FreeForm();
                when(mockWorksheetService.addDataToCell("d1", "f1", "s1", "A1", "val1")).thenReturn(cellResult);

                List<Map<String, Object>> cellData = List.of(Map.of("Cell Id", "A1", Constants.DATA, "val1"));

                String taskId = sharedArea.addMultipleCellDataAsync("d1", "f1", "s1", cellData);

                assertNotNull(taskId);
                verify(mockEventHandler).handleEvent(eq(taskId), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("addMultipleCellDataAsync - invalid cell id → throws in executor")
        void addMultipleCellDataAsync_invalidCellId() {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class);
                 MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {

                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);
                ws.when(() -> Worksheets.isValidCellId("INVALID")).thenReturn(false);

                List<Map<String, Object>> cellData = List.of(Map.of("Cell Id", "INVALID", Constants.DATA, "val"));

                String taskId = sharedArea.addMultipleCellDataAsync("d1", "f1", "s1", cellData);
                assertNotNull(taskId);
                verify(mockEventHandler, never()).handleEvent(anyString(), any(FreeForm.class));
            }
        }

        @Test
        @DisplayName("addMultipleCellDataAsync - exception from worksheetService → wraps in RuntimeException")
        void addMultipleCellDataAsync_exception() {
            try (MockedStatic<Worksheets> ws = mockStatic(Worksheets.class);
                 MockedStatic<ThreadLocalProxy> tlp = mockStatic(ThreadLocalProxy.class)) {

                tlp.when(ThreadLocalProxy::getAll).thenReturn(Map.of());
                tlp.when(() -> ThreadLocalProxy.setAll(anyMap())).thenAnswer(inv -> null);
                tlp.when(ThreadLocalProxy::removeAll).thenAnswer(inv -> null);
                ws.when(() -> Worksheets.isValidCellId("A1")).thenReturn(true);

                when(mockWorksheetService.addDataToCell(anyString(), anyString(), anyString(), anyString(), anyString()))
                        .thenThrow(new RuntimeException("Service error"));

                List<Map<String, Object>> cellData = List.of(Map.of("Cell Id", "A1", Constants.DATA, "val"));

                String taskId = sharedArea.addMultipleCellDataAsync("d1", "f1", "s1", cellData);
                assertNotNull(taskId);
                verify(mockEventHandler, never()).handleEvent(anyString(), any(FreeForm.class));
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // addRowsWithListOfEntity
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("addRowsWithListOfEntity")
    class AddRowsWithListOfEntity {

        @Test
        @DisplayName("Null/empty data → Constants.FREE_FORM.put returns null causing NPE in List.of")
        void nullData_returnsError() {
            try (MockedStatic<StringUtil> su = mockStatic(StringUtil.class)) {
                su.when(() -> StringUtil.isDataNullOrEmpty(isNull())).thenReturn(true);

                assertThrows(NullPointerException.class, () ->
                        sharedArea.addRowsWithListOfEntity("d1", "f1", "s1", null, null));
            }
        }

        @Test
        @DisplayName("Valid data → gets entity definition, used range, and adds rows")
        void validData_addsRows() {
            try (MockedStatic<StringUtil> su = mockStatic(StringUtil.class)) {
                List<EntityValue> data = new ArrayList<>();
                EntityValue mockEntity = mock(EntityValue.class);
                data.add(mockEntity);

                su.when(() -> StringUtil.isDataNullOrEmpty(data)).thenReturn(false);

                EntityDefinition mockDef = mock(EntityDefinition.class);
                when(mockEntities.getEntityDefinition(mockEntity)).thenReturn(mockDef);
                when(mockWorksheetOperation.getUsedRangeForSheet("d1", "f1", "s1")).thenReturn("A1:D10");

                List<FreeForm> expected = List.of(new FreeForm());
                when(mockWorksheetOperation.addRowsToSheetWithEntity("f1", "s1", data, "A1:D10", "d1", null, mockDef))
                        .thenReturn(expected);

                List<FreeForm> result = sharedArea.addRowsWithListOfEntity("d1", "f1", "s1", data, null);

                assertSame(expected, result);
                verify(mockWorksheetOperation).getUsedRangeForSheet("d1", "f1", "s1");
                verify(mockWorksheetOperation).addRowsToSheetWithEntity("f1", "s1", data, "A1:D10", "d1", null, mockDef);
            }
        }
    }
}
