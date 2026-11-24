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
import app.krista.model.base.File;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("SheetOperationsFacade Tests")
class SheetOperationsFacadeTest {

    @Mock
    private CellResponseService cellResponseService;

    @Mock
    private WorksheetResponseService worksheetResponseService;

    @Mock
    private RowResponseService rowResponseService;

    @Mock
    private ColumnResponseService columnResponseService;

    @Mock
    private FormulaResponseService formulaResponseService;

    @Mock
    private SearchResponseService searchResponseService;

    @Mock
    private WorkbookResponseService workbookResponseService;

    @Mock
    private PdfResponseService pdfResponseService;

    @Mock
    private ExtensionResponse mockResponse;

    private SheetOperationsFacade facade;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        facade = new SheetOperationsFacade(cellResponseService, worksheetResponseService, rowResponseService,
                columnResponseService, formulaResponseService, searchResponseService,
                workbookResponseService, pdfResponseService);
    }

    // Cell Operations

    @Test
    @DisplayName("getGetCellValueExtensionResponse delegates to CellResponseService")
    void testGetGetCellValueExtensionResponse() {
        String cellId = "A1";
        when(cellResponseService.getGetCellValueExtensionResponse(FILE_ID, SHEET_NAME, cellId, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getGetCellValueExtensionResponse(FILE_ID, SHEET_NAME, cellId, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(cellResponseService).getGetCellValueExtensionResponse(FILE_ID, SHEET_NAME, cellId, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getAddDataToCellExtensionResponse delegates to CellResponseService")
    void testGetAddDataToCellExtensionResponse() {
        String cellId = "B2";
        String data = "testData";
        when(cellResponseService.getAddDataToCellExtensionResponse(FILE_ID, SHEET_NAME, cellId, data, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddDataToCellExtensionResponse(FILE_ID, SHEET_NAME, cellId, data, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(cellResponseService).getAddDataToCellExtensionResponse(FILE_ID, SHEET_NAME, cellId, data, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getAddMultipleCellDataExtensionResponse delegates to CellResponseService")
    void testGetAddMultipleCellDataExtensionResponse() {
        List<Map<String, Object>> cellIDData = List.of(Map.of("cellId", "A1", "data", "val1"));
        when(cellResponseService.getAddMultipleCellDataExtensionResponse(FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddMultipleCellDataExtensionResponse(FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(cellResponseService).getAddMultipleCellDataExtensionResponse(FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getAddFormulaToCellExtensionResponse delegates to CellResponseService")
    void testGetAddFormulaToCellExtensionResponse() {
        String cellId = "C3";
        String formula = "=SUM(A1:A10)";
        when(cellResponseService.getAddFormulaToCellExtensionResponse(FILE_ID, SHEET_NAME, cellId, formula, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddFormulaToCellExtensionResponse(FILE_ID, SHEET_NAME, cellId, formula, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(cellResponseService).getAddFormulaToCellExtensionResponse(FILE_ID, SHEET_NAME, cellId, formula, ALLOW_RETRY);
    }

    // Worksheet Operations

    @Test
    @DisplayName("getDeleteSheetExtensionResponse delegates to WorksheetResponseService")
    void testGetDeleteSheetExtensionResponse() {
        when(worksheetResponseService.getDeleteSheetExtensionResponse(FILE_ID, SHEET_NAME, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getDeleteSheetExtensionResponse(FILE_ID, SHEET_NAME, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(worksheetResponseService).getDeleteSheetExtensionResponse(FILE_ID, SHEET_NAME, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getCreateSheetExtensionResponse delegates to WorksheetResponseService")
    void testGetCreateSheetExtensionResponse() {
        String newSheetName = "NewSheet";
        when(worksheetResponseService.getCreateSheetExtensionResponse(FILE_ID, newSheetName, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getCreateSheetExtensionResponse(FILE_ID, newSheetName, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(worksheetResponseService).getCreateSheetExtensionResponse(FILE_ID, newSheetName, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getGetSheetsInformationExtensionResponse delegates to WorksheetResponseService")
    void testGetGetSheetsInformationExtensionResponse() {
        when(worksheetResponseService.getGetSheetsInformationExtensionResponse(FILE_ID, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getGetSheetsInformationExtensionResponse(FILE_ID, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(worksheetResponseService).getGetSheetsInformationExtensionResponse(FILE_ID, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getClearSheetForRangeExtensionResponse delegates to WorksheetResponseService")
    void testGetClearSheetForRangeExtensionResponse() {
        String cellRange = "A1:Z100";
        when(worksheetResponseService.getClearSheetForRangeExtensionResponse(FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getClearSheetForRangeExtensionResponse(FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(worksheetResponseService).getClearSheetForRangeExtensionResponse(FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY);
    }

    // Row Operations

    @Test
    @DisplayName("getAddRowsExtensionResponse delegates to RowResponseService")
    void testGetAddRowsExtensionResponse() {
        List<String> rows = List.of("row1", "row2");
        List<File> rowsFromCsv = null;
        when(rowResponseService.getAddRowsExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddRowsExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(rowResponseService).getAddRowsExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getAddRowWithIndexExtensionResponse delegates to RowResponseService")
    void testGetAddRowWithIndexExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsv = null;
        Double rowIndex = 3.0;
        when(rowResponseService.getAddRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(rowResponseService).getAddRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getGetCompleteRowInfoExtensionResponse delegates to RowResponseService")
    void testGetGetCompleteRowInfoExtensionResponse() {
        Double rowIndex = 5.0;
        when(rowResponseService.getGetCompleteRowInfoExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getGetCompleteRowInfoExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(rowResponseService).getGetCompleteRowInfoExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getGetRowsForRangeExtensionResponse delegates to RowResponseService")
    void testGetGetRowsForRangeExtensionResponse() {
        String range = "A1:C5";
        when(rowResponseService.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(rowResponseService).getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getUpdateRowWithIndexExtensionResponse delegates to RowResponseService")
    void testGetUpdateRowWithIndexExtensionResponse() {
        String row = "updatedRow";
        Double rowIndex = 2.0;
        when(rowResponseService.getUpdateRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getUpdateRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(rowResponseService).getUpdateRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getDeleteRowWithIndexExtensionResponse delegates to RowResponseService")
    void testGetDeleteRowWithIndexExtensionResponse() {
        Double rowIndex = 4.0;
        when(rowResponseService.getDeleteRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getDeleteRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(rowResponseService).getDeleteRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);
    }

    // Column Operations

    @Test
    @DisplayName("getDeleteColumnWithIndexExtensionResponse delegates to ColumnResponseService")
    void testGetDeleteColumnWithIndexExtensionResponse() {
        String columnName = "B";
        when(columnResponseService.getDeleteColumnWithIndexExtensionResponse(FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getDeleteColumnWithIndexExtensionResponse(FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(columnResponseService).getDeleteColumnWithIndexExtensionResponse(FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getGetColumnValuesForKeysInSheetExtensionResponse delegates to ColumnResponseService")
    void testGetGetColumnValuesForKeysInSheetExtensionResponse() {
        String range = "A1:D10";
        List<String> keyRows = List.of("key1");
        List<String> keyColumns = List.of("col1");
        Boolean getAllRecords = false;
        Double columnNumber = 1.0;
        when(columnResponseService.getGetColumnValuesForKeysInSheetExtensionResponse(FILE_ID, SHEET_NAME, range, keyRows, keyColumns, getAllRecords, columnNumber, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getGetColumnValuesForKeysInSheetExtensionResponse(FILE_ID, SHEET_NAME, range, keyRows, keyColumns, getAllRecords, columnNumber, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(columnResponseService).getGetColumnValuesForKeysInSheetExtensionResponse(FILE_ID, SHEET_NAME, range, keyRows, keyColumns, getAllRecords, columnNumber, ALLOW_RETRY);
    }

    // Formula Operations

    @Test
    @DisplayName("getExecuteFormulaExtensionResponse delegates to FormulaResponseService")
    void testGetExecuteFormulaExtensionResponse() {
        String formula = "=SUM(A1:A10)";
        String formulaName = "SUM";
        String comment = "test";
        when(formulaResponseService.getExecuteFormulaExtensionResponse(FILE_ID, formula, formulaName, comment, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getExecuteFormulaExtensionResponse(FILE_ID, formula, formulaName, comment, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(formulaResponseService).getExecuteFormulaExtensionResponse(FILE_ID, formula, formulaName, comment, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getVlookupExtensionResponse delegates to FormulaResponseService")
    void testGetVlookupExtensionResponse() {
        String lookupValue = "test";
        String range = "A1:B10";
        Double columnIndex = 2.0;
        Boolean exactMatch = true;
        when(formulaResponseService.getVlookupExtensionResponse(FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getVlookupExtensionResponse(FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(formulaResponseService).getVlookupExtensionResponse(FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getSummationExtensionResponse delegates to FormulaResponseService")
    void testGetSummationExtensionResponse() {
        String range = "A1:A10";
        when(formulaResponseService.getSummationExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getSummationExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(formulaResponseService).getSummationExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY);
    }

    // Search Operations

    @Test
    @DisplayName("getFindValueExtensionResponse delegates to SearchResponseService")
    void testGetFindValueExtensionResponse() {
        String lookupValue = "findMe";
        when(searchResponseService.getFindValueExtensionResponse(FILE_ID, SHEET_NAME, lookupValue, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getFindValueExtensionResponse(FILE_ID, SHEET_NAME, lookupValue, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(searchResponseService).getFindValueExtensionResponse(FILE_ID, SHEET_NAME, lookupValue, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse delegates to SearchResponseService")
    void testGetGetMinMaxValueFromRowDataByColumnRangeExtensionResponse() {
        String keyColumnName = "Amount";
        String columnRange = "A1:A100";
        when(searchResponseService.getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(FILE_ID, SHEET_NAME, keyColumnName, columnRange, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(FILE_ID, SHEET_NAME, keyColumnName, columnRange, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(searchResponseService).getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(FILE_ID, SHEET_NAME, keyColumnName, columnRange, ALLOW_RETRY);
    }

    // Workbook Operations

    @Test
    @DisplayName("getGetExcelFileInformationExtensionResponse delegates to WorkbookResponseService")
    void testGetGetExcelFileInformationExtensionResponse() {
        String path = "/documents";
        String name = "test.xlsx";
        when(workbookResponseService.getGetExcelFileInformationExtensionResponse(path, name, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getGetExcelFileInformationExtensionResponse(path, name, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(workbookResponseService).getGetExcelFileInformationExtensionResponse(path, name, ALLOW_RETRY);
    }

    // PDF Operations

    @Test
    @DisplayName("getUploadFileAndConvertToPDFExtensionResponse delegates to PdfResponseService")
    void testGetUploadFileAndConvertToPDFExtensionResponse() {
        File excelFile = mock(File.class);
        String filePath = "/path/to/file.xlsx";
        when(pdfResponseService.getUploadFileAndConvertToPDFExtensionResponse(excelFile, filePath, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getUploadFileAndConvertToPDFExtensionResponse(excelFile, filePath, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(pdfResponseService).getUploadFileAndConvertToPDFExtensionResponse(excelFile, filePath, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getConvertToPDFExtensionResponse delegates to PdfResponseService")
    void testGetConvertToPDFExtensionResponse() {
        when(pdfResponseService.getConvertToPDFExtensionResponse(FILE_ID, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getConvertToPDFExtensionResponse(FILE_ID, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(pdfResponseService).getConvertToPDFExtensionResponse(FILE_ID, ALLOW_RETRY);
    }
}
