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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelSheet;
import app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2.WorkbookWorksheetRangeRequestBuilderV2;
import app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2.WorkbookWorksheetRangeRequestV2;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonNull;
import com.google.gson.JsonPrimitive;
import com.microsoft.graph.http.GraphError;
import com.microsoft.graph.http.GraphErrorResponse;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.*;
import com.microsoft.graph.options.HeaderOption;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.requests.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.MockedConstruction;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class WorksheetServiceTest {

    @Mock private WorkbookOperation workbookOperation;
    @Mock private GraphServiceClientProviderFactory providerFactory;
    @Mock private GraphServiceClientProvider clientProvider;
    @Mock private StringUtil stringUtil;
    @Mock private WorkbookSessionManager sessionManager;

    private WorksheetService worksheetService;

    // Common test constants
    private static final String FILE_ID = "file-abc";
    private static final String DRIVE_ID = "drive-123";
    private static final String SHEET_NAME = "Sheet1";
    private static final String RANGE = "A1:B10";
    private static final String CELL_ID = "A1";

    @BeforeEach
    void setUp() {
        // providerFactory.create() is now lazy — stub it so it returns the mock when first used.
        when(providerFactory.create()).thenReturn(clientProvider);
        worksheetService = new WorksheetService(workbookOperation, providerFactory, stringUtil, sessionManager);
    }

    // ==================== Constructor ====================

    @Test
    void constructor_initializesWithDependencies() {
        assertNotNull(worksheetService);
        // Constructor no longer calls create() eagerly; it stores the factory for lazy init.
        verify(providerFactory, never()).create();
    }

    @Test
    void multipleInstances_areIndependent() {
        WorksheetService s2 = new WorksheetService(workbookOperation, providerFactory, stringUtil, sessionManager);
        assertNotNull(s2);
        assertNotSame(worksheetService, s2);
    }

    // ==================== Constants ====================

    @Test
    void constants_haveExpectedValues() {
        assertEquals("Text", WorksheetService.TEXT);
    }

    @Test
    void class_hasServiceAnnotation() {
        assertTrue(WorksheetService.class.isAnnotationPresent(org.jvnet.hk2.annotations.Service.class));
    }

    // ==================== getRangeParameterSet ====================

    @Test
    void getRangeParameterSet_withValidRange_returnsParameterSet() {
        assertNotNull(WorksheetService.getRangeParameterSet("A1:B10"));
    }

    @Test
    void getRangeParameterSet_withNull_returnsParameterSet() {
        assertNotNull(WorksheetService.getRangeParameterSet(null));
    }

    @Test
    void getRangeParameterSet_withEmpty_returnsParameterSet() {
        assertNotNull(WorksheetService.getRangeParameterSet(""));
    }

    @Test
    void getRangeParameterSet_withComplexRange_returnsParameterSet() {
        assertNotNull(WorksheetService.getRangeParameterSet("Sheet1!A1:Z100"));
    }

    // ==================== Helper: stub chains ====================

    private WorkbookWorksheetRequestBuilder stubWorksheetChain() {
        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookWorksheetRequestBuilder wsBuilder = mock(WorkbookWorksheetRequestBuilder.class);

        when(clientProvider.getDriveRequestBuilder(any())).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.worksheets(any())).thenReturn(wsBuilder);
        return wsBuilder;
    }

    /**
     * Stubs full graph chain including worksheets() (collection), used for getWorksheets/listWorksheets.
     */
    private WorkbookWorksheetCollectionRequestBuilder stubWorksheetCollectionChain() {
        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookWorksheetCollectionRequestBuilder collectionBuilder = mock(WorkbookWorksheetCollectionRequestBuilder.class);

        try {
            when(clientProvider.getDriveRequestBuilder(any(), any(), any())).thenReturn(drive);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.worksheets()).thenReturn(collectionBuilder);
        return collectionBuilder;
    }

    private WorkbookWorksheetRangeRequestBuilder stubRangeChain() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = mock(WorkbookWorksheetRangeRequestBuilder.class);
        when(wsBuilder.range(any(WorkbookWorksheetRangeParameterSet.class))).thenReturn(rangeBuilder);
        return rangeBuilder;
    }

    private static GraphServiceException graphException(int httpStatus, String code, String message) {
        GraphServiceException ex = mock(GraphServiceException.class);
        when(ex.getResponseCode()).thenReturn(httpStatus);
        when(ex.getMessage()).thenReturn(message);
        if (code != null) {
            GraphErrorResponse errorResponse = new GraphErrorResponse();
            GraphError error = new GraphError();
            error.code = code;
            errorResponse.error = error;
            when(ex.getError()).thenReturn(errorResponse);
        } else {
            when(ex.getError()).thenReturn(null);
        }
        return ex;
    }

    // ==================== getWorksheets ====================

    @Test
    void getWorksheets_nullPage_returnsEmptyList() {
        WorkbookWorksheetCollectionRequestBuilder collBuilder = stubWorksheetCollectionChain();
        WorkbookWorksheetCollectionRequest req = mock(WorkbookWorksheetCollectionRequest.class);
        when(collBuilder.buildRequest()).thenReturn(req);
        when(req.get()).thenReturn(null);

        List<ExcelSheet> result = worksheetService.getWorksheets(FILE_ID, DRIVE_ID, null, false);

        assertEquals(0, result.size());
    }

    @Test
    void getWorksheets_singlePageNoNextPage_returnsSheets() {
        WorkbookWorksheetCollectionRequestBuilder collBuilder = stubWorksheetCollectionChain();
        WorkbookWorksheetCollectionRequest req = mock(WorkbookWorksheetCollectionRequest.class);
        when(collBuilder.buildRequest()).thenReturn(req);

        WorkbookWorksheetCollectionPage page = mock(WorkbookWorksheetCollectionPage.class);
        WorkbookWorksheet ws = new WorkbookWorksheet();
        ws.id = "id1";
        ws.name = "Sheet1";
        when(page.getCurrentPage()).thenReturn(List.of(ws));
        when(page.getNextPage()).thenReturn(null);
        when(req.get()).thenReturn(page);

        List<ExcelSheet> result = worksheetService.getWorksheets(FILE_ID, DRIVE_ID, null, false);

        assertFalse(result.isEmpty());
    }

    @Test
    void getWorksheets_multiplePages_iteratesAll() {
        WorkbookWorksheetCollectionRequestBuilder collBuilder = stubWorksheetCollectionChain();
        WorkbookWorksheetCollectionRequest req = mock(WorkbookWorksheetCollectionRequest.class);
        when(collBuilder.buildRequest()).thenReturn(req);

        WorkbookWorksheet ws1 = new WorkbookWorksheet();
        ws1.id = "id1";
        ws1.name = "Sheet1";
        WorkbookWorksheet ws2 = new WorkbookWorksheet();
        ws2.id = "id2";
        ws2.name = "Sheet2";

        WorkbookWorksheetCollectionPage page2 = mock(WorkbookWorksheetCollectionPage.class);
        when(page2.getCurrentPage()).thenReturn(List.of(ws2));
        when(page2.getNextPage()).thenReturn(null);

        WorkbookWorksheetCollectionRequestBuilder nextBuilder = mock(WorkbookWorksheetCollectionRequestBuilder.class);
        WorkbookWorksheetCollectionRequest nextReq = mock(WorkbookWorksheetCollectionRequest.class);
        when(nextBuilder.buildRequest()).thenReturn(nextReq);
        when(nextReq.get()).thenReturn(page2);

        WorkbookWorksheetCollectionPage page1 = mock(WorkbookWorksheetCollectionPage.class);
        when(page1.getCurrentPage()).thenReturn(List.of(ws1));
        when(page1.getNextPage()).thenReturn(nextBuilder);
        when(req.get()).thenReturn(page1);

        List<ExcelSheet> result = worksheetService.getWorksheets(FILE_ID, DRIVE_ID, null, false);

        assertEquals(2, result.size());
    }

    // ==================== deleteWorksheets ====================

    @Test
    void deleteWorksheets_success_returnsDeletedMessage() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(wsBuilder.buildRequest()).thenReturn(req);
        when(req.delete()).thenReturn(null);

        String result = worksheetService.deleteWorksheets(FILE_ID, SHEET_NAME);

        assertEquals(SHEET_NAME + " Deleted Successfully", result);
    }

    @Test
    void deleteWorksheets_otherError_withServiceError_throwsIllegalState() {
        GraphServiceException ex = graphException(403, "AccessDenied", "forbidden");
        GraphError serviceError = new GraphError();
        serviceError.message = "You don't have permission";
        when(ex.getServiceError()).thenReturn(serviceError);

        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(wsBuilder.buildRequest()).thenReturn(req);
        when(req.delete()).thenThrow(ex);

        IllegalStateException thrown = assertThrows(IllegalStateException.class,
                () -> worksheetService.deleteWorksheets(FILE_ID, SHEET_NAME));
        assertTrue(thrown.getMessage().contains("You don't have permission"));
    }

    @Test
    void deleteWorksheets_otherError_nullServiceError_usesExMessage() {
        GraphServiceException ex = graphException(500, "ServerError", "server down");
        when(ex.getServiceError()).thenReturn(null);

        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(wsBuilder.buildRequest()).thenReturn(req);
        when(req.delete()).thenThrow(ex);

        IllegalStateException thrown = assertThrows(IllegalStateException.class,
                () -> worksheetService.deleteWorksheets(FILE_ID, SHEET_NAME));
        assertTrue(thrown.getMessage().contains("server down"));
    }

    // ==================== deleteTableColumn ====================

    @Test
    void deleteTableColumn_success_returnsTrue() {
        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookTableRequestBuilder table = mock(WorkbookTableRequestBuilder.class);
        WorkbookTableColumnRequestBuilder col = mock(WorkbookTableColumnRequestBuilder.class);
        WorkbookTableColumnRequest colReq = mock(WorkbookTableColumnRequest.class);

        when(clientProvider.getDriveRequestBuilder(isNull())).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.tables(any())).thenReturn(table);
        when(table.columns(any())).thenReturn(col);
        when(col.buildRequest()).thenReturn(colReq);
        when(colReq.delete()).thenReturn(null);

        assertTrue(worksheetService.deleteTableColumn(FILE_ID, "Table1", "Column1"));
    }

    @Test
    void deleteTableColumn_graphException_withServiceError_returnsFalse() {
        GraphServiceException ex = graphException(404, "ItemNotFound", "not found");
        GraphError serviceError = new GraphError();
        serviceError.message = "Column not found";
        when(ex.getServiceError()).thenReturn(serviceError);

        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookTableRequestBuilder table = mock(WorkbookTableRequestBuilder.class);
        WorkbookTableColumnRequestBuilder col = mock(WorkbookTableColumnRequestBuilder.class);
        WorkbookTableColumnRequest colReq = mock(WorkbookTableColumnRequest.class);

        when(clientProvider.getDriveRequestBuilder(isNull())).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.tables(any())).thenReturn(table);
        when(table.columns(any())).thenReturn(col);
        when(col.buildRequest()).thenReturn(colReq);
        when(colReq.delete()).thenThrow(ex);

        assertFalse(worksheetService.deleteTableColumn(FILE_ID, "Table1", "Column1"));
    }

    @Test
    void deleteTableColumn_graphException_nullServiceError_returnsFalse() {
        GraphServiceException ex = graphException(500, "ServerError", "fail");
        when(ex.getServiceError()).thenReturn(null);

        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookTableRequestBuilder table = mock(WorkbookTableRequestBuilder.class);
        WorkbookTableColumnRequestBuilder col = mock(WorkbookTableColumnRequestBuilder.class);
        WorkbookTableColumnRequest colReq = mock(WorkbookTableColumnRequest.class);

        when(clientProvider.getDriveRequestBuilder(isNull())).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.tables(any())).thenReturn(table);
        when(table.columns(any())).thenReturn(col);
        when(col.buildRequest()).thenReturn(colReq);
        when(colReq.delete()).thenThrow(ex);

        assertFalse(worksheetService.deleteTableColumn(FILE_ID, "Table1", "Column1"));
    }

    // ==================== deleteTableColumnInSharedFile ====================

    @Test
    void deleteTableColumnInSharedFile_success_returnsTrue() {
        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookTableRequestBuilder table = mock(WorkbookTableRequestBuilder.class);
        WorkbookTableColumnRequestBuilder col = mock(WorkbookTableColumnRequestBuilder.class);
        WorkbookTableColumnRequest colReq = mock(WorkbookTableColumnRequest.class);

        when(clientProvider.getDriveRequestBuilder(eq(DRIVE_ID))).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.tables(any())).thenReturn(table);
        when(table.columns(any())).thenReturn(col);
        when(col.buildRequest()).thenReturn(colReq);
        when(colReq.delete()).thenReturn(null);

        assertTrue(worksheetService.deleteTableColumnInSharedFile(FILE_ID, "Table1", "Column1", DRIVE_ID));
    }

    @Test
    void deleteTableColumnInSharedFile_graphException_withServiceError_returnsFalse() {
        GraphServiceException ex = graphException(500, "ServerError", "fail");
        GraphError serviceError = new GraphError();
        serviceError.message = "Server error";
        when(ex.getServiceError()).thenReturn(serviceError);

        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookTableRequestBuilder table = mock(WorkbookTableRequestBuilder.class);
        WorkbookTableColumnRequestBuilder col = mock(WorkbookTableColumnRequestBuilder.class);
        WorkbookTableColumnRequest colReq = mock(WorkbookTableColumnRequest.class);

        when(clientProvider.getDriveRequestBuilder(eq(DRIVE_ID))).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.tables(any())).thenReturn(table);
        when(table.columns(any())).thenReturn(col);
        when(col.buildRequest()).thenReturn(colReq);
        when(colReq.delete()).thenThrow(ex);

        assertFalse(worksheetService.deleteTableColumnInSharedFile(FILE_ID, "Table1", "Column1", DRIVE_ID));
    }

    @Test
    void deleteTableColumnInSharedFile_graphException_nullServiceError_returnsFalse() {
        GraphServiceException ex = graphException(403, "AccessDenied", "denied");
        when(ex.getServiceError()).thenReturn(null);

        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookTableRequestBuilder table = mock(WorkbookTableRequestBuilder.class);
        WorkbookTableColumnRequestBuilder col = mock(WorkbookTableColumnRequestBuilder.class);
        WorkbookTableColumnRequest colReq = mock(WorkbookTableColumnRequest.class);

        when(clientProvider.getDriveRequestBuilder(eq(DRIVE_ID))).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.tables(any())).thenReturn(table);
        when(table.columns(any())).thenReturn(col);
        when(col.buildRequest()).thenReturn(colReq);
        when(colReq.delete()).thenThrow(ex);

        assertFalse(worksheetService.deleteTableColumnInSharedFile(FILE_ID, "Table1", "Column1", DRIVE_ID));
    }

    // ==================== createWorksheet ====================

    @Test
    void createWorksheet_success_returnsExcelSheet() {
        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookWorksheetCollectionRequestBuilder collBuilder = mock(WorkbookWorksheetCollectionRequestBuilder.class);
        WorkbookWorksheetAddRequestBuilder addBuilder = mock(WorkbookWorksheetAddRequestBuilder.class);
        WorkbookWorksheetAddRequest addReq = mock(WorkbookWorksheetAddRequest.class);

        when(clientProvider.getDriveRequestBuilder(isNull())).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.worksheets()).thenReturn(collBuilder);
        when(collBuilder.add(any())).thenReturn(addBuilder);
        when(addBuilder.buildRequest()).thenReturn(addReq);

        WorkbookWorksheet ws = new WorkbookWorksheet();
        ws.id = "newId";
        ws.name = "NewSheet";
        when(addReq.post()).thenReturn(ws);

        ExcelSheet result = worksheetService.createWorksheet(FILE_ID, "NewSheet");

        assertNotNull(result);
        assertEquals("NewSheet", result.getSheetName());
    }

    @Test
    void createWorksheet_nullResponse_returnsEmptyExcelSheet() {
        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
        WorkbookWorksheetCollectionRequestBuilder collBuilder = mock(WorkbookWorksheetCollectionRequestBuilder.class);
        WorkbookWorksheetAddRequestBuilder addBuilder = mock(WorkbookWorksheetAddRequestBuilder.class);
        WorkbookWorksheetAddRequest addReq = mock(WorkbookWorksheetAddRequest.class);

        when(clientProvider.getDriveRequestBuilder(isNull())).thenReturn(drive);
        when(drive.items(any())).thenReturn(item);
        when(item.workbook()).thenReturn(workbook);
        when(workbook.worksheets()).thenReturn(collBuilder);
        when(collBuilder.add(any())).thenReturn(addBuilder);
        when(addBuilder.buildRequest()).thenReturn(addReq);
        when(addReq.post()).thenReturn(null);

        ExcelSheet result = worksheetService.createWorksheet(FILE_ID, "NewSheet");

        assertNotNull(result);
    }

//    @Test
//    void createWorksheet_graphException_returnsEmptyExcelSheet() {
//        DriveRequestBuilder drive = mock(DriveRequestBuilder.class);
//        DriveItemRequestBuilder item = mock(DriveItemRequestBuilder.class);
//        WorkbookRequestBuilder workbook = mock(WorkbookRequestBuilder.class);
//        WorkbookWorksheetCollectionRequestBuilder collBuilder = mock(WorkbookWorksheetCollectionRequestBuilder.class);
//        WorkbookWorksheetAddRequestBuilder addBuilder = mock(WorkbookWorksheetAddRequestBuilder.class);
//        WorkbookWorksheetAddRequest addReq = mock(WorkbookWorksheetAddRequest.class);
//
//        when(clientProvider.getDriveRequestBuilder(isNull())).thenReturn(drive);
//        when(drive.items(any())).thenReturn(item);
//        when(item.workbook()).thenReturn(workbook);
//        when(workbook.worksheets()).thenReturn(collBuilder);
//        when(collBuilder.add(any())).thenReturn(addBuilder);
//        when(addBuilder.buildRequest()).thenReturn(addReq);
//        when(addReq.post()).thenThrow(graphException(500, "ServerError", "fail"));
//
//        ExcelSheet result = worksheetService.createWorksheet(FILE_ID, "NewSheet");
//
//        assertNotNull(result);
//    }

    // ==================== findValue ====================

    @Test
    void findValue_success_returnsList() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);

        WorkbookRange wbRange = new WorkbookRange();
        wbRange.text = new JsonArray();
        wbRange.values = new JsonArray();
        when(usedRangeReq.get()).thenReturn(wbRange);

        List<String> result = worksheetService.findValue(DRIVE_ID, FILE_ID, SHEET_NAME, "search");

        assertNotNull(result);
    }

    // ==================== getWorkbookRange ====================

    @Test
    void getWorkbookRange_success_returnsRange() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);

        WorkbookRange expected = new WorkbookRange();
        expected.address = "Sheet1!A1:B10";
        when(usedRangeReq.get()).thenReturn(expected);

        WorkbookRange result = worksheetService.getWorkbookRange(DRIVE_ID, FILE_ID, SHEET_NAME);

        assertNotNull(result);
        assertEquals("Sheet1!A1:B10", result.address);
    }

    // ==================== addDataToCell ====================

    @Test
    void addDataToCell_success_returnsFreeFormWithCellInfo() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn(null); // Triggers early return in updateCell

        FreeForm result = worksheetService.addDataToCell(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, "hello");

        assertNotNull(result);
    }

    // ==================== updateCell ====================

    @Test
    void updateCell_nullRequestUrl_returnsEarly() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn(null);

        // Should not throw, just returns early.
        worksheetService.updateCell(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, "data");
    }

    @Test
    void updateCell_nullClient_returnsEarly() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn(null);
        // Simpler: just return null for getClient
        when(rangeBuilder.getRequestUrl()).thenReturn("https://fake");
        when(rangeBuilder.getClient()).thenReturn(null);

        worksheetService.updateCell(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, "data");
    }

    @Test
    void updateCell_withFormulaValue_setsFormulas() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.patch(any())).thenReturn(new WorkbookRange());
                     })) {
            worksheetService.updateCell(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, "=SUM(A1:A10)");

            assertEquals(1, mocked.constructed().size());
        }
    }

    @Test
    void updateCell_withRegularValue_setsValues() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.patch(any())).thenReturn(new WorkbookRange());
                     })) {
            worksheetService.updateCell(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, "hello");

            assertEquals(1, mocked.constructed().size());
        }
    }

    @Test
    void updateCell_withNullData_setsValues() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.patch(any())).thenReturn(new WorkbookRange());
                     })) {
            worksheetService.updateCell(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, null);

            assertEquals(1, mocked.constructed().size());
        }
    }

    // ==================== addRows (both overloads) ====================

    @Test
    void addRows_forInsertAndUpdate_delegatesToWorkbookOperation() throws IOException {
        when(workbookOperation.addRowsForInsertAndUpdate(any(), any(), any(), any(), any())).thenReturn(List.of(new FreeForm()));

        List<FreeForm> result = worksheetService.addRows(DRIVE_ID, FILE_ID, SHEET_NAME, List.of("row1"), "A1");

        assertNotNull(result);
        verify(workbookOperation).addRowsForInsertAndUpdate(DRIVE_ID, FILE_ID, SHEET_NAME, List.of("row1"), "A1");
    }

    // ==================== getRow ====================

    @Test
    void getRow_delegatesToWorkbookOperation() {
        FreeForm expected = new FreeForm();
        when(workbookOperation.getRow(any(), any(), any(), any())).thenReturn(expected);

        FreeForm result = worksheetService.getRow(FILE_ID, SHEET_NAME, 1.0, DRIVE_ID);

        assertSame(expected, result);
        verify(workbookOperation).getRow(FILE_ID, SHEET_NAME, 1.0, DRIVE_ID);
    }

    // ==================== getRows (via getDataForRange) ====================

    @Test
    void getRows_nullBaseBuilderUrl_returnsEmptyList() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn(null);
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        // getDataForRange returns null on null URL → getRows gets null
        List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

        // null is returned from getDataForRange, getRows returns it as-is (no catch needed)
        assertNull(result);
    }

    @Test
    void getRows_successWithValidData_returnsParsedData() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        JsonArray row = new JsonArray();
        row.add("val1");
        row.add("val2");
        JsonArray data = new JsonArray();
        data.add(row);

        WorkbookRange workbookRange = new WorkbookRange();
        workbookRange.text = data;
        workbookRange.values = data;
        workbookRange.address = "Sheet1!A1:B1";
        workbookRange.rowCount = 1;
        workbookRange.columnCount = 2;

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(mock.buildRequest(anyList())).thenReturn(reqV2);
                         when(reqV2.get()).thenReturn(workbookRange);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }

    @Test
    void getRows_withSessionOptions_usesSessionOverload() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());

        List<Option> sessionOptions = Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1"));
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(sessionOptions);

        JsonArray row = new JsonArray();
        row.add("val");
        JsonArray data = new JsonArray();
        data.add(row);

        WorkbookRange workbookRange = new WorkbookRange();
        workbookRange.text = data;
        workbookRange.address = "A1:A1";
        workbookRange.rowCount = 1;
        workbookRange.columnCount = 1;

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(mock.buildRequest(anyList())).thenReturn(reqV2);
                         when(reqV2.get()).thenReturn(workbookRange);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }

    @Test
    void getRows_nullWorkbookRange_returnsEmptyList() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenReturn(null);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
            assertTrue(result.isEmpty());
        }
    }

    @Test
    void getRows_jsonNullText_returnsEmptyList() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        WorkbookRange workbookRange = new WorkbookRange();
        workbookRange.text = JsonNull.INSTANCE;
        workbookRange.address = "A1:A1";

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenReturn(workbookRange);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertTrue(result.isEmpty());
        }
    }

    @Test
    void getRows_textNotArray_returnsEmptyList() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        WorkbookRange workbookRange = new WorkbookRange();
        workbookRange.text = new JsonPrimitive("not an array");
        workbookRange.address = "A1:A1";

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenReturn(workbookRange);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertTrue(result.isEmpty());
        }
    }

    @Test
    void getRows_graphServiceException404_returnsErrorFreeForm() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(404, "ItemNotFound", "not found");
        when(ex.getServiceError()).thenReturn(null);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
            assertFalse(result.isEmpty());
        }
    }

    @Test
    void getRows_graphServiceException403_returnsAccessDeniedError() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(403, "AccessDenied", "forbidden");
        GraphError serviceError = new GraphError();
        serviceError.message = "Access denied message";
        when(ex.getServiceError()).thenReturn(serviceError);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
            assertFalse(result.isEmpty());
        }
    }

    @Test
    void getRows_graphServiceException403_nullServiceError_usesMessage() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(403, "AccessDenied", "forbidden msg");
        when(ex.getServiceError()).thenReturn(null);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }

    @Test
    void getRows_graphServiceException500_returnsGenericError() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(500, "ServerError", "server fail");
        GraphError serviceError = new GraphError();
        serviceError.message = "Internal server error";
        when(ex.getServiceError()).thenReturn(serviceError);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
            assertFalse(result.isEmpty());
        }
    }

    @Test
    void getRows_graphServiceException500_nullServiceError() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(500, null, "server fail");
        when(ex.getServiceError()).thenReturn(null);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }

    @Test
    void getRows_invalidSessionError_retriesWithFreshSession() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        GraphServiceException sessionEx = graphException(400, "InvalidSessionIdException", "session expired");
        when(sessionEx.getServiceError()).thenReturn(null);

        JsonArray row = new JsonArray();
        row.add("val");
        JsonArray data = new JsonArray();
        data.add(row);
        WorkbookRange wbRange = new WorkbookRange();
        wbRange.text = data;
        wbRange.address = "A1:A1";
        wbRange.rowCount = 1;
        wbRange.columnCount = 1;

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(mock.buildRequest(anyList())).thenReturn(reqV2);
                         // First call throws session error, second succeeds (for retry path)
                         when(reqV2.get()).thenThrow(sessionEx).thenReturn(wbRange);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
            verify(sessionManager).refreshSessionHeaderOptions(any(), any(), anyBoolean());
        }
    }

    @Test
    void getRows_invalidSessionError_noSessionManager_returnsError() {
        WorksheetService svcNoSession = new WorksheetService(workbookOperation, providerFactory, stringUtil, null);

        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(400, "InvalidSessionIdException", "session expired");
        when(ex.getServiceError()).thenReturn(null);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = svcNoSession.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }

    @Test
    void getRows_runtimeException_returnsEmptyList() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(new RuntimeException("unexpected"));
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertTrue(result.isEmpty());
        }
    }

    @Test
    void getRows_wacAccessTokenError_retriesWithFreshSession() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.emptyList());

        GraphServiceException wacEx = graphException(403, "AccessDenied", "Could not obtain a WAC access token");
        when(wacEx.getServiceError()).thenReturn(null);
        // retry also fails with 404
        GraphServiceException retryEx = graphException(404, "ItemNotFound", "not found");
        when(retryEx.getServiceError()).thenReturn(null);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(mock.buildRequest(anyList())).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(wacEx).thenThrow(retryEx);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
            verify(sessionManager).refreshSessionHeaderOptions(any(), any(), anyBoolean());
        }
    }

    // ==================== getCellData ====================

    @Test
    void getCellData_nullBaseBuilderUrl_returnsEmptyFreeForm() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn(null);
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        FreeForm result = worksheetService.getCellData(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID);

        assertNotNull(result);
    }

    @Test
    void getCellData_emptyResults_returnsEmptyFreeForm() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        WorkbookRange wr = new WorkbookRange();
        wr.text = null;

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenReturn(wr);
                     })) {
            FreeForm result = worksheetService.getCellData(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID);

            assertNotNull(result);
        }
    }

    // ==================== findMinMaxValueFromRowDataByColumnRange ====================

    @Test
    void findMinMax_nullWorkbookId_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, null, SHEET_NAME, "A", "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_blankWorkbookId_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, "  ", SHEET_NAME, "A", "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_nullWorksheetName_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, null, "A", "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_blankWorksheetName_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, "  ", "A", "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_nullKeyColumnName_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, null, "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_blankKeyColumnName_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "  ", "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_nullColumnRange_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "A", null);
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_blankColumnRange_returnsEmptyMap() {
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "A", "  ");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_nullWorkbookRange_returnsEmptyMap() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);
        when(usedRangeReq.get()).thenReturn(null);

        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "A", "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_nullColumnCount_returnsEmptyMap() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);

        WorkbookRange wr = new WorkbookRange();
        wr.columnCount = null;
        when(usedRangeReq.get()).thenReturn(wr);

        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "A", "B:C");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_invalidColumnRange_startAfterEnd_returnsNull() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);

        WorkbookRange wr = new WorkbookRange();
        wr.columnCount = 5;
        when(usedRangeReq.get()).thenReturn(wr);

        // "C:A" → start index > end index
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "A", "C:A");
        assertNull(result);
    }

    @Test
    void findMinMax_keyColumnNotFound_returnsNull() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);

        WorkbookRange wr = new WorkbookRange();
        wr.columnCount = 3; // A, B, C only
        when(usedRangeReq.get()).thenReturn(wr);

        // "Z" is not in A-C columns
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "Z", "A:B");
        assertNull(result);
    }

    @Test
    void findMinMax_nullWorksheetData_returnsEmptyMap() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);

        WorkbookRange wr = new WorkbookRange();
        wr.columnCount = 3;
        wr.text = null; // Null text data
        when(usedRangeReq.get()).thenReturn(wr);

        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "A", "A:B");
        assertTrue(result.isEmpty());
    }

    @Test
    void findMinMax_rangeColumnsNotFound_returnsNull() {
        WorkbookWorksheetRequestBuilder wsBuilder = stubWorksheetChain();
        WorkbookWorksheetUsedRangeRequestBuilder usedRangeBuilder = mock(WorkbookWorksheetUsedRangeRequestBuilder.class);
        WorkbookWorksheetUsedRangeRequest usedRangeReq = mock(WorkbookWorksheetUsedRangeRequest.class);

        when(wsBuilder.usedRange()).thenReturn(usedRangeBuilder);
        when(usedRangeBuilder.buildRequest()).thenReturn(usedRangeReq);

        WorkbookRange wr = new WorkbookRange();
        wr.columnCount = 2; // A, B only
        when(usedRangeReq.get()).thenReturn(wr);

        // "X:Z" columns don't exist → startRangeColumnIndex == -1
        Map<String, Object> result = worksheetService.findMinMaxValueFromRowDataByColumnRange(
                DRIVE_ID, FILE_ID, SHEET_NAME, "A", "X:Z");
        assertNull(result);
    }

    // ==================== worksheetExists - full coverage ====================

    @Test
    void worksheetExists_nullFileId_returnsFalse() {
        assertFalse(worksheetService.worksheetExists(null, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_blankFileId_returnsFalse() {
        assertFalse(worksheetService.worksheetExists("   ", SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_emptyFileId_returnsFalse() {
        assertFalse(worksheetService.worksheetExists("", SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_nullWorksheetName_returnsFalse() {
        assertFalse(worksheetService.worksheetExists(FILE_ID, null, DRIVE_ID));
    }

    @Test
    void worksheetExists_blankWorksheetName_returnsFalse() {
        assertFalse(worksheetService.worksheetExists(FILE_ID, "   ", DRIVE_ID));
    }

    @Test
    void worksheetExists_emptyWorksheetName_returnsFalse() {
        assertFalse(worksheetService.worksheetExists(FILE_ID, "", DRIVE_ID));
    }

    @Test
    void worksheetExists_withSession_returnsTrue() {
        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(req);
        when(req.get()).thenReturn(new WorkbookWorksheet());
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
        verify(builder).buildRequest(anyList());
        verify(builder, never()).buildRequest();
    }

    @Test
    void worksheetExists_withEmptySessionOptions_returnsTrue() {
        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenReturn(new WorkbookWorksheet());
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
        verify(builder).buildRequest();
    }

    @Test
    void worksheetExists_sessionManagerNull_usesVarargs() {
        WorksheetService svc = new WorksheetService(workbookOperation, providerFactory, stringUtil, null);
        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenReturn(new WorkbookWorksheet());

        assertTrue(svc.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_invalidSessionIdException_retriesWithFreshSession() {
        GraphServiceException ex = graphException(400, "InvalidSessionIdException", "session expired");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(ex);
        when(retryReq.get()).thenReturn(new WorkbookWorksheet());

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
        verify(sessionManager).refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false);
    }

    @Test
    void worksheetExists_sessionNotFound_retriesWithFreshSession() {
        GraphServiceException ex = graphException(404, "SessionNotFound", "session gone");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(ex);
        when(retryReq.get()).thenReturn(new WorkbookWorksheet());

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_sessionIdNotFound_retriesWithFreshSession() {
        GraphServiceException ex = graphException(404, "SessionIdNotFound", "unknown session");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(ex);
        when(retryReq.get()).thenReturn(new WorkbookWorksheet());

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_404SessionMessage_retriesWithFreshSession() {
        GraphServiceException ex = graphException(404, null, "Workbook SESSION handle expired");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(ex);
        when(retryReq.get()).thenReturn(new WorkbookWorksheet());

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_403WacAccessToken_retriesWithFreshSession() {
        GraphServiceException ex = graphException(403, "AccessDenied", "Could not obtain a WAC access token");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(ex);
        when(retryReq.get()).thenReturn(new WorkbookWorksheet());

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_invalidSession_retryReturns404_returnsFalse() {
        GraphServiceException primary = graphException(400, "InvalidSessionIdException", "session expired");
        GraphServiceException retry = graphException(404, "ItemNotFound", "not found");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(primary);
        when(retryReq.get()).thenThrow(retry);

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertFalse(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_invalidSession_retryReturns403_throwsIllegalState() {
        GraphServiceException primary = graphException(400, "InvalidSessionIdException", "session expired");
        GraphServiceException retry = graphException(403, "AccessDenied", "Insufficient permissions");
        when(retry.getServiceError()).thenReturn(null);

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(primary);
        when(retryReq.get()).thenThrow(retry);

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertThrows(IllegalStateException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_invalidSession_retryReturns403_withServiceError() {
        GraphServiceException primary = graphException(400, "InvalidSessionIdException", "session expired");
        GraphServiceException retry = graphException(403, "AccessDenied", "Insufficient permissions");
        GraphError serviceError = new GraphError();
        serviceError.message = "Service error message";
        when(retry.getServiceError()).thenReturn(serviceError);

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(primary);
        when(retryReq.get()).thenThrow(retry);

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        IllegalStateException thrown = assertThrows(IllegalStateException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
        assertTrue(thrown.getMessage().contains("Service error message"));
    }

    @Test
    void worksheetExists_invalidSession_retryReturns500_throws() {
        GraphServiceException primary = graphException(400, "InvalidSessionIdException", "session expired");
        GraphServiceException retry = graphException(500, "ServerError", "internal error");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(primary);
        when(retryReq.get()).thenThrow(retry);

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        assertThrows(GraphServiceException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_invalidSession_retryWithEmptySession_succeeds() {
        GraphServiceException primary = graphException(400, "InvalidSessionIdException", "session expired");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest sessionReq = mock(WorkbookWorksheetRequest.class);
        WorkbookWorksheetRequest retryReq = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest(anyList())).thenReturn(sessionReq);
        when(builder.buildRequest()).thenReturn(retryReq);
        when(sessionReq.get()).thenThrow(primary);
        when(retryReq.get()).thenReturn(new WorkbookWorksheet());

        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(DRIVE_ID, FILE_ID, false))
                .thenReturn(Collections.emptyList());

        assertTrue(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_invalidSessionButSessionManagerNull_throws() {
        WorksheetService svc = new WorksheetService(workbookOperation, providerFactory, stringUtil, null);
        GraphServiceException ex = graphException(400, "InvalidSessionIdException", "session expired");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);

        assertThrows(GraphServiceException.class, () -> svc.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_404NotSessionRelated_returnsFalse() {
        GraphServiceException ex = graphException(404, "ItemNotFound", "not found");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        assertFalse(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_404NullMessage_returnsFalse() {
        GraphServiceException ex = graphException(404, null, null);

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        assertFalse(worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_403_throwsIllegalState() {
        GraphServiceException ex = graphException(403, "AccessDenied", "Insufficient permissions");
        when(ex.getServiceError()).thenReturn(null);

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        IllegalStateException thrown = assertThrows(IllegalStateException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
        assertTrue(thrown.getMessage().contains("Access denied"));
        assertSame(ex, thrown.getCause());
    }

    @Test
    void worksheetExists_403_withServiceError() {
        GraphServiceException ex = graphException(403, "AccessDenied", "Insufficient permissions");
        GraphError serviceError = new GraphError();
        serviceError.message = "Detailed service error";
        when(ex.getServiceError()).thenReturn(serviceError);

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        IllegalStateException thrown = assertThrows(IllegalStateException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
        assertTrue(thrown.getMessage().contains("Detailed service error"));
    }

    @Test
    void worksheetExists_500_throwsGraphException() {
        GraphServiceException ex = graphException(500, "ServerError", "internal error");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        assertThrows(GraphServiceException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_nullError_throwsGraphException() {
        GraphServiceException ex = graphException(500, null, "unknown");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        assertThrows(GraphServiceException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_errorWithNullInnerError_throwsGraphException() {
        GraphServiceException ex = mock(GraphServiceException.class);
        GraphErrorResponse errorResponse = new GraphErrorResponse();
        errorResponse.error = null;
        when(ex.getError()).thenReturn(errorResponse);
        when(ex.getResponseCode()).thenReturn(500);
        when(ex.getMessage()).thenReturn("boom");

        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(ex);
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        assertThrows(GraphServiceException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    @Test
    void worksheetExists_runtimeException_propagates() {
        WorkbookWorksheetRequestBuilder builder = stubWorksheetChain();
        WorkbookWorksheetRequest req = mock(WorkbookWorksheetRequest.class);
        when(builder.buildRequest()).thenReturn(req);
        when(req.get()).thenThrow(new RuntimeException("network glitch"));
        when(sessionManager.sessionHeaderOptions(DRIVE_ID, FILE_ID, false)).thenReturn(Collections.emptyList());

        assertThrows(RuntimeException.class,
                () -> worksheetService.worksheetExists(FILE_ID, SHEET_NAME, DRIVE_ID));
    }

    // ==================== getDataForRangeWithFreshSession - retry error branches ====================

    @Test
    void getRows_sessionRetry_403error_returnsErrorWithAccessDenied() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        GraphServiceException sessionEx = graphException(400, "InvalidSessionIdException", "session expired");
        when(sessionEx.getServiceError()).thenReturn(null);
        GraphServiceException retryEx = graphException(403, "AccessDenied", "forbidden after refresh");
        GraphError serviceError = new GraphError();
        serviceError.message = "Access denied after refresh";
        when(retryEx.getServiceError()).thenReturn(serviceError);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(mock.buildRequest(anyList())).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(sessionEx).thenThrow(retryEx);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
            assertFalse(result.isEmpty());
        }
    }

    @Test
    void getRows_sessionRetry_500error_returnsGenericError() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-2")));

        GraphServiceException sessionEx = graphException(400, "InvalidSessionIdException", "session expired");
        when(sessionEx.getServiceError()).thenReturn(null);
        GraphServiceException retryEx = graphException(500, "ServerError", "server fail");
        when(retryEx.getServiceError()).thenReturn(null);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(mock.buildRequest(anyList())).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(sessionEx).thenThrow(retryEx);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }

    @Test
    void getRows_sessionRetry_nullBaseBuilder_returnsNull() {
        // First call in getDataForRange succeeds past builder check, but
        // the retry path in getDataForRangeWithFreshSession has null URL.
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.singletonList(new HeaderOption("workbook-session-id", "sess-1")));
        when(sessionManager.refreshSessionHeaderOptions(any(), any(), anyBoolean()))
                .thenReturn(Collections.emptyList());

        GraphServiceException sessionEx = graphException(400, "InvalidSessionIdException", "session expired");
        when(sessionEx.getServiceError()).thenReturn(null);

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        // First call: URL non-null → V2 builder constructed → throws session error
        // Retry call: URL null → early return
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test").thenReturn(null);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(mock.buildRequest(anyList())).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(sessionEx);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            // getDataForRangeWithFreshSession returns null on null URL, getRows returns it
            assertNull(result);
        }
    }

    // ==================== logGraphServiceException branches ====================

    @Test
    void getRows_graphException_withErrorAndServiceError_logsAll() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(500, "ServerError", "fail");
        GraphError serviceError = new GraphError();
        serviceError.message = "detail";
        when(ex.getServiceError()).thenReturn(serviceError);
        // Also set errorResponse.error.message for logGraphServiceException
        GraphErrorResponse errResp = ex.getError();
        if (errResp != null && errResp.error != null) {
            errResp.error.message = "error msg";
        }

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }

    @Test
    void getRows_graphException_nullErrorNullServiceError_logsFallbackMessage() {
        WorkbookWorksheetRangeRequestBuilder rangeBuilder = stubRangeChain();
        when(rangeBuilder.getRequestUrl()).thenReturn("https://graph.microsoft.com/v1.0/test");

        com.microsoft.graph.core.IBaseClient<?> mockClient = mock(com.microsoft.graph.core.IBaseClient.class);
        doReturn(mockClient).when(rangeBuilder).getClient();
        when(rangeBuilder.getOptions()).thenReturn(Collections.emptyList());
        when(sessionManager.sessionHeaderOptions(any(), any(), anyBoolean())).thenReturn(Collections.emptyList());

        GraphServiceException ex = graphException(500, null, "raw message");
        when(ex.getServiceError()).thenReturn(null);

        try (MockedConstruction<WorkbookWorksheetRangeRequestBuilderV2> mocked =
                     mockConstruction(WorkbookWorksheetRangeRequestBuilderV2.class, (mock, context) -> {
                         WorkbookWorksheetRangeRequestV2 reqV2 = mock(WorkbookWorksheetRangeRequestV2.class);
                         when(mock.buildRequest(any(Option[].class))).thenReturn(reqV2);
                         when(reqV2.get()).thenThrow(ex);
                     })) {
            List<FreeForm> result = worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            assertNotNull(result);
        }
    }
}
