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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedCellResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedFileResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedSheetResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedTableResponseService;
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
@DisplayName("SharedOperationsFacade Tests")
class SharedOperationsFacadeTest {

    @Mock
    private SharedCellResponseService sharedCellResponseService;

    @Mock
    private SharedFileResponseService sharedFileResponseService;

    @Mock
    private SharedSheetResponseService sharedSheetResponseService;

    @Mock
    private SharedTableResponseService sharedTableResponseService;

    private SharedOperationsFacade facade;

    @Mock
    private ExtensionResponse mockResponse;

    private static final String DRIVE_ID = "driveId";
    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final String TABLE_NAME = "Table1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        facade = new SharedOperationsFacade(sharedCellResponseService, sharedFileResponseService,
                sharedSheetResponseService, sharedTableResponseService);
    }

    // Shared Cell Operations

    @Test
    @DisplayName("getAddDataToCellExtensionResponse delegates to SharedCellResponseService")
    void testGetAddDataToCellExtensionResponse() {
        String cellId = "A1";
        String data = "testData";
        when(sharedCellResponseService.getAddDataToCellExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellId, data, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddDataToCellExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellId, data, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedCellResponseService).getAddDataToCellExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellId, data, ALLOW_RETRY);
        verifyNoInteractions(sharedFileResponseService, sharedSheetResponseService, sharedTableResponseService);
    }

    @Test
    @DisplayName("getCellDataExtensionResponse delegates to SharedCellResponseService")
    void testGetCellDataExtensionResponse() {
        String cellId = "B2";
        when(sharedCellResponseService.getCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellId, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellId, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedCellResponseService).getCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellId, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getAddMultipleCellDataExtensionResponse delegates to SharedCellResponseService")
    void testGetAddMultipleCellDataExtensionResponse() {
        List<Map<String, Object>> cellIDData = List.of(Map.of("cellId", "A1", "data", "val1"));
        when(sharedCellResponseService.getAddMultipleCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddMultipleCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedCellResponseService).getAddMultipleCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);
    }

    // Shared File Operations

    @Test
    @DisplayName("getSharedExcelFileExtensionResponse delegates to SharedFileResponseService")
    void testGetSharedExcelFileExtensionResponse() {
        String folderName = "SharedFolder";
        String fileName = "test.xlsx";
        when(sharedFileResponseService.getSharedExcelFileExtensionResponse(folderName, fileName, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getSharedExcelFileExtensionResponse(folderName, fileName, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedFileResponseService).getSharedExcelFileExtensionResponse(folderName, fileName, ALLOW_RETRY);
        verifyNoInteractions(sharedCellResponseService, sharedSheetResponseService, sharedTableResponseService);
    }

    @Test
    @DisplayName("getConvertToPdfExtensionResponse delegates to SharedFileResponseService")
    void testGetConvertToPdfExtensionResponse() {
        when(sharedFileResponseService.getConvertToPdfExtensionResponse(DRIVE_ID, FILE_ID, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getConvertToPdfExtensionResponse(DRIVE_ID, FILE_ID, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedFileResponseService).getConvertToPdfExtensionResponse(DRIVE_ID, FILE_ID, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getExecuteFormulaExtensionResponse delegates to SharedFileResponseService")
    void testGetExecuteFormulaExtensionResponse() {
        String formula = "=SUM(A1:A10)";
        String formulaName = "SUM";
        String comment = "test comment";
        when(sharedFileResponseService.getExecuteFormulaExtensionResponse(DRIVE_ID, FILE_ID, formula, formulaName, comment, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getExecuteFormulaExtensionResponse(DRIVE_ID, FILE_ID, formula, formulaName, comment, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedFileResponseService).getExecuteFormulaExtensionResponse(DRIVE_ID, FILE_ID, formula, formulaName, comment, ALLOW_RETRY);
    }

    // Shared Sheet Operations

    @Test
    @DisplayName("getSheetExtensionResponse delegates to SharedSheetResponseService")
    void testGetSheetExtensionResponse() {
        List<String> rows = List.of("row1", "row2");
        List<File> rowsFromFile = null;
        when(sharedSheetResponseService.getSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromFile, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromFile, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromFile, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getSheetVlookupExtensionResponse delegates to SharedSheetResponseService")
    void testGetSheetVlookupExtensionResponse() {
        String lookupValue = "test";
        String range = "A1:B10";
        Double columnIndex = 2.0;
        Boolean exactMatch = true;
        when(sharedSheetResponseService.getSheetVlookupExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getSheetVlookupExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getSheetVlookupExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getRowsForRangeExtensionResponse delegates to SharedSheetResponseService")
    void testGetRowsForRangeExtensionResponse() {
        String range = "A1:C5";
        when(sharedSheetResponseService.getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getSearchExtensionResponse delegates to SharedSheetResponseService")
    void testGetSearchExtensionResponse() {
        String dataToSearch = "findMe";
        when(sharedSheetResponseService.getSearchExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, dataToSearch, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getSearchExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, dataToSearch, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getSearchExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, dataToSearch, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getAddRowWithIndexExtensionResponse delegates to SharedSheetResponseService")
    void testGetAddRowWithIndexExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsv = null;
        Double rowIndex = 3.0;
        when(sharedSheetResponseService.getAddRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getAddRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getUpdateRowWithIndexExtensionResponse delegates to SharedSheetResponseService")
    void testGetUpdateRowWithIndexExtensionResponse() {
        String row = "updatedRow";
        Double rowIndex = 2.0;
        when(sharedSheetResponseService.getUpdateRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getUpdateRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getUpdateRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getColumnValuesForKeysInSheetExtensionResponse delegates to SharedSheetResponseService")
    void testGetColumnValuesForKeysInSheetExtensionResponse() {
        String range = "A1:D10";
        List<String> lookupRows = List.of("key1");
        List<String> lookupColumns = List.of("col1");
        Boolean getAllRecords = false;
        Double columnNumber = 1.0;
        when(sharedSheetResponseService.getColumnValuesForKeysInSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, lookupRows, lookupColumns, getAllRecords, columnNumber, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getColumnValuesForKeysInSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, lookupRows, lookupColumns, getAllRecords, columnNumber, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getColumnValuesForKeysInSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, lookupRows, lookupColumns, getAllRecords, columnNumber, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getClearSharedSheetForRangeExtensionResponse delegates to SharedSheetResponseService")
    void testGetClearSharedSheetForRangeExtensionResponse() {
        String cellRange = "A1:Z100";
        when(sharedSheetResponseService.getClearSharedSheetForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getClearSharedSheetForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getClearSharedSheetForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getRangeSumExtensionResponse delegates to SharedSheetResponseService")
    void testGetRangeSumExtensionResponse() {
        String range = "A1:A10";
        when(sharedSheetResponseService.getRangeSumExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getRangeSumExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getRangeSumExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getDeleteColumnWithIndexExtensionResponse delegates to SharedSheetResponseService")
    void testGetDeleteColumnWithIndexExtensionResponse() {
        String columnName = "B";
        when(sharedSheetResponseService.getDeleteColumnWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getDeleteColumnWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getDeleteColumnWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getCompleteRowInfoExtensionResponse delegates to SharedSheetResponseService")
    void testGetCompleteRowInfoExtensionResponse() {
        Double rowIndex = 5.0;
        when(sharedSheetResponseService.getCompleteRowInfoExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getCompleteRowInfoExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getCompleteRowInfoExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getDeleteRowWithIndexExtensionResponse delegates to SharedSheetResponseService")
    void testGetDeleteRowWithIndexExtensionResponse() {
        Double rowIndex = 3.0;
        when(sharedSheetResponseService.getDeleteRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getDeleteRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedSheetResponseService).getDeleteRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);
    }

    // Shared Table Operations

    @Test
    @DisplayName("getExtensionResponse delegates to SharedTableResponseService")
    void testGetExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsvFile = null;
        when(sharedTableResponseService.getExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, rows, rowsFromCsvFile, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, rows, rowsFromCsvFile, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedTableResponseService).getExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, rows, rowsFromCsvFile, ALLOW_RETRY);
        verifyNoInteractions(sharedCellResponseService, sharedFileResponseService, sharedSheetResponseService);
    }

    @Test
    @DisplayName("getTableRowsExtensionResponse delegates to SharedTableResponseService")
    void testGetTableRowsExtensionResponse() {
        when(sharedTableResponseService.getTableRowsExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getTableRowsExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedTableResponseService).getTableRowsExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getColumnValuesForKeysExtensionResponse delegates to SharedTableResponseService")
    void testGetColumnValuesForKeysExtensionResponse() {
        List<String> keyRows = List.of("key1");
        List<String> lookupColumn = List.of("col1");
        Boolean getAllRecord = true;
        when(sharedTableResponseService.getColumnValuesForKeysExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, keyRows, lookupColumn, getAllRecord, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getColumnValuesForKeysExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, keyRows, lookupColumn, getAllRecord, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedTableResponseService).getColumnValuesForKeysExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, keyRows, lookupColumn, getAllRecord, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getDeleteColumnExtensionResponse delegates to SharedTableResponseService")
    void testGetDeleteColumnExtensionResponse() {
        String columnName = "ColumnA";
        when(sharedTableResponseService.getDeleteColumnExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getDeleteColumnExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedTableResponseService).getDeleteColumnExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getVlookupInTableExtensionResponse delegates to SharedTableResponseService")
    void testGetVlookupInTableExtensionResponse() {
        String lookupValue = "value";
        Double columnIndex = 2.0;
        Boolean exactMatch = false;
        when(sharedTableResponseService.getVlookupInTableExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getVlookupInTableExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedTableResponseService).getVlookupInTableExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getColumnsSumExtensionResponse delegates to SharedTableResponseService")
    void testGetColumnsSumExtensionResponse() {
        List<String> columns = List.of("col1", "col2");
        when(sharedTableResponseService.getColumnsSumExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columns, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getColumnsSumExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(sharedTableResponseService).getColumnsSumExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);
    }
}
