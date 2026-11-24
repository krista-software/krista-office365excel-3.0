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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.impl.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static org.junit.jupiter.api.Assertions.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("ReentryHandlerFactory Tests")
class ReentryHandlerFactoryTest {

    @Mock
    private ErrorHandlingStateManager stateManager;

    @Mock
    private ExtensionResponseGenerator responseGenerator;

    @Mock
    private WorkbookService workbookService;

    @Mock
    private WorksheetService worksheetService;

    @Mock
    private TableService tableService;

    @Mock
    private RestClientProvider restClientProvider;

    @Mock
    private DriveService driveService;

    @Mock
    private WorksheetOperation worksheetOperation;

    @Mock
    private StringUtil stringUtil;

    @Mock
    private SheetFormattingService sheetFormattingService;

    @Mock
    private TelemetryHelper telemetryHelper;

    private ReentryHandlerFactory factory;

    @BeforeEach
    void setUp() {
        factory = new ReentryHandlerFactory(stateManager, responseGenerator, workbookService,
                worksheetService, tableService, restClientProvider, driveService,
                worksheetOperation, stringUtil, sheetFormattingService, telemetryHelper);
    }

    @Test
    @DisplayName("getTableReentryHandler returns TableReentryHandlerImpl instance")
    void testGetTableReentryHandler() {
        TableReentryHandler handler = factory.getTableReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(TableReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getSheetReentryHandler returns SheetReentryHandlerImpl instance")
    void testGetSheetReentryHandler() {
        SheetReentryHandler handler = factory.getSheetReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(SheetReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getRangeReentryHandler returns RangeReentryHandlerImpl instance")
    void testGetRangeReentryHandler() {
        RangeReentryHandler handler = factory.getRangeReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(RangeReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getCellReentryHandler returns CellReentryHandlerImpl instance")
    void testGetCellReentryHandler() {
        CellReentryHandler handler = factory.getCellReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(CellReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getSharedFileReentryHandler returns SharedFileReentryHandlerImpl instance")
    void testGetSharedFileReentryHandler() {
        SharedFileReentryHandler handler = factory.getSharedFileReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(SharedFileReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getColumnReentryHandler returns ColumnReentryHandlerImpl instance")
    void testGetColumnReentryHandler() {
        ColumnReentryHandler handler = factory.getColumnReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(ColumnReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getPdfReentryHandler returns PdfReentryHandlerImpl instance")
    void testGetPdfReentryHandler() {
        PdfReentryHandler handler = factory.getPdfReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(PdfReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getWorksheetReentryHandler returns WorksheetReentryHandlerImpl instance")
    void testGetWorksheetReentryHandler() {
        WorksheetReentryHandler handler = factory.getWorksheetReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(WorksheetReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getRowReentryHandler returns RowReentryHandlerImpl instance")
    void testGetRowReentryHandler() {
        RowReentryHandler handler = factory.getRowReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(RowReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getFindValueReentryHandler returns FindValueReentryHandlerImpl instance")
    void testGetFindValueReentryHandler() {
        FindValueReentryHandler handler = factory.getFindValueReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(FindValueReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getWorkbookReentryHandler returns WorkbookReentryHandlerImpl instance")
    void testGetWorkbookReentryHandler() {
        WorkbookReentryHandler handler = factory.getWorkbookReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(WorkbookReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getFormulaReentryHandler returns FormulaReentryHandlerImpl instance")
    void testGetFormulaReentryHandler() {
        FormulaReentryHandler handler = factory.getFormulaReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(FormulaReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getSummationReentryHandler returns SummationReentryHandlerImpl instance")
    void testGetSummationReentryHandler() {
        SummationReentryHandler handler = factory.getSummationReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(SummationReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getColumnValuesReentryHandler returns ColumnValuesReentryHandlerImpl instance")
    void testGetColumnValuesReentryHandler() {
        ColumnValuesReentryHandler handler = factory.getColumnValuesReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(ColumnValuesReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getFileReentryHandler returns FileReentryHandlerImpl instance")
    void testGetFileReentryHandler() {
        FileReentryHandler handler = factory.getFileReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(FileReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getMinMaxValueReentryHandler returns MinMaxValueReentryHandlerImpl instance")
    void testGetMinMaxValueReentryHandler() {
        MinMaxValueReentryHandler handler = factory.getMinMaxValueReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(MinMaxValueReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getTableVlookupReentryHandler returns TableVlookupReentryHandlerImpl instance")
    void testGetTableVlookupReentryHandler() {
        TableVlookupReentryHandler handler = factory.getTableVlookupReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(TableVlookupReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getTableRowsReentryHandler returns TableRowsReentryHandlerImpl instance")
    void testGetTableRowsReentryHandler() {
        TableRowsReentryHandler handler = factory.getTableRowsReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(TableRowsReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getTableColumnReentryHandler returns DeleteColumnReentryHandlerImpl instance")
    void testGetTableColumnReentryHandler() {
        DeleteColumnReentryHandler handler = factory.getTableColumnReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(DeleteColumnReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("getTableCreationReentryHandler returns TableCreationReentryHandlerImpl instance")
    void testGetTableCreationReentryHandler() {
        TableCreationReentryHandler handler = factory.getTableCreationReentryHandler();

        assertNotNull(handler);
        assertInstanceOf(TableCreationReentryHandlerImpl.class, handler);
    }

    @Test
    @DisplayName("Each factory method returns a new instance on each call")
    void testFactoryMethodsReturnNewInstances() {
        TableReentryHandler handler1 = factory.getTableReentryHandler();
        TableReentryHandler handler2 = factory.getTableReentryHandler();

        assertNotNull(handler1);
        assertNotNull(handler2);
        assertNotSame(handler1, handler2);
    }

    @Test
    @DisplayName("Different factory methods return different handler types")
    void testDifferentFactoryMethodsReturnDifferentTypes() {
        Object tableHandler = factory.getTableReentryHandler();
        Object sheetHandler = factory.getSheetReentryHandler();
        Object cellHandler = factory.getCellReentryHandler();
        Object rangeHandler = factory.getRangeReentryHandler();
        Object pdfHandler = factory.getPdfReentryHandler();

        assertNotSame(tableHandler.getClass(), sheetHandler.getClass());
        assertNotSame(sheetHandler.getClass(), cellHandler.getClass());
        assertNotSame(cellHandler.getClass(), rangeHandler.getClass());
        assertNotSame(rangeHandler.getClass(), pdfHandler.getClass());
    }
}
