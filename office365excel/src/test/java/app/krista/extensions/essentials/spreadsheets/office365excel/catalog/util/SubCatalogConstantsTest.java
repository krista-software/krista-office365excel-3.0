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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

import static org.junit.jupiter.api.Assertions.*;

class SubCatalogConstantsTest {

    @Test
    void testValidationResults() {
        assertEquals("VALIDATION_RESULTS", SubCatalogConstants.VALIDATION_RESULTS);
    }

    @Test
    void testReenter() {
        assertEquals("REENTER", SubCatalogConstants.REENTER);
    }

    @Test
    void testConfirmReenterAddRowsToTable() {
        assertEquals("confirmReenterAddRowsToTable", SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_TABLE);
    }

    @Test
    void testHandleReenterAddRowsToTable() {
        assertEquals("handleReenterAddRowsToTable", SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_TABLE);
    }

    @Test
    void testConfirmReenterAddRowsToSheet() {
        assertEquals("confirmReenterAddRowsToSheet", SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_SHEET);
    }

    @Test
    void testHandleReenterAddRowsToSheet() {
        assertEquals("handleReenterAddRowsToSheet", SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_SHEET);
    }

    @Test
    void testConfirmReenterRangeSum() {
        assertEquals("confirmReenterRangeSum", SubCatalogConstants.CONFIRM_REENTER_RANGE_SUM);
    }

    @Test
    void testHandleReenterRangeSum() {
        assertEquals("handleReenterRangeSum", SubCatalogConstants.HANDLE_REENTER_RANGE_SUM);
    }

    @Test
    void testConfirmReenterGetTableRows() {
        assertEquals("confirmReenterGetTableRows", SubCatalogConstants.CONFIRM_REENTER_GET_TABLE_ROWS);
    }

    @Test
    void testHandleReenterGetTableRows() {
        assertEquals("handleReenterGetTableRows", SubCatalogConstants.HANDLE_REENTER_GET_TABLE_ROWS);
    }

    @Test
    void testConfirmReenterGetRowsForRange() {
        assertEquals("confirmReenterGetRowsForRange", SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE);
    }

    @Test
    void testHandleReenterGetRowsForRange() {
        assertEquals("handleReenterGetRowsForRange", SubCatalogConstants.HANDLE_REENTER_GET_ROWS_FOR_RANGE);
    }

    @Test
    void testConfirmReenterGetRowsForRangeForSheet() {
        assertEquals("confirmReenterGetRowsForRangeForSheet", SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET);
    }

    @Test
    void testHandleReenterGetRowsForRangeForSheet() {
        assertEquals("handleReenterGetRowsForRangeForSheet", SubCatalogConstants.HANDLE_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET);
    }

    @Test
    void testSheetVlookupConstants() {
        assertEquals("confirmReenterSheetVlookup", SubCatalogConstants.CONFIRM_REENTER_SHARED_SHEET_VLOOKUP);
        assertEquals("handleReenterSheetVlookup", SubCatalogConstants.HANDLE_REENTER_SHARED_SHEET_VLOOKUP);
        assertEquals("confirmReenterSheetVlookup", SubCatalogConstants.CONFIRM_REENTER_SHEET_VLOOKUP);
        assertEquals("handleReenterSheetVlookup", SubCatalogConstants.HANDLE_REENTER_SHEET_VLOOKUP);
    }

    @Test
    void testAddDataToCellConstants() {
        assertEquals("confirmReenterAddDataToCell", SubCatalogConstants.CONFIRM_REENTER_ADD_DATA_TO_CELL);
        assertEquals("handleReenterAddDataToCell", SubCatalogConstants.HANDLE_REENTER_ADD_DATA_TO_CELL);
    }

    @Test
    void testVlookupInTableConstants() {
        assertEquals("confirmReenterVlookupInTable", SubCatalogConstants.CONFIRM_REENTER_VLOOKUP_IN_TABLE);
        assertEquals("handleReenterVlookupInTable", SubCatalogConstants.HANDLE_REENTER_VLOOKUP_IN_TABLE);
    }

    @Test
    void testDeleteRowWithIndexConstants() {
        assertEquals("confirmReenterDeleteRowWithIndex", SubCatalogConstants.CONFIRM_REENTER_DELETE_ROW_WITH_INDEX);
        assertEquals("handleReenterDeleteRowWithIndex", SubCatalogConstants.HANDLE_REENTER_DELETE_ROW_WITH_INDEX);
        assertEquals("confirmReenterDeleteRowWithIndex", SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET_ROW_WITH_INDEX);
        assertEquals("handleReenterDeleteRowWithIndex", SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET_ROW_WITH_INDEX);
    }

    @Test
    void testGetCompleteRowInfoConstants() {
        assertEquals("confirmReenterGetCompleteRowInfo", SubCatalogConstants.CONFIRM_REENTER_GET_COMPLETE_ROW_INFO);
        assertEquals("handleReenterGetCompleteRowInfo", SubCatalogConstants.HANDLE_REENTER_GET_COMPLETE_ROW_INFO);
        assertEquals("confirmReenterGetCompleteRowInfo", SubCatalogConstants.CONFIRM_REENTER_GET_COMPLETE_SHEET_ROW_INFO);
        assertEquals("handleReenterGetCompleteRowInfo", SubCatalogConstants.HANDLE_REENTER_GET_COMPLETE_SHEET_ROW_INFO);
    }

    @Test
    void testDeleteColumnConstants() {
        assertEquals("confirmReenterDeleteColumn", SubCatalogConstants.CONFIRM_REENTER_DELETE_COLUMN);
        assertEquals("handleReenterDeleteColumn", SubCatalogConstants.HANDLE_REENTER_DELETE_COLUMN);
    }

    @Test
    void testExecuteFormulaConstants() {
        assertEquals("confirmReenterExecuteFormula", SubCatalogConstants.CONFIRM_REENTER_EXECUTE_FORMULA);
        assertEquals("handleReenterExecuteFormula", SubCatalogConstants.HANDLE_REENTER_EXECUTE_FORMULA);
    }

    @Test
    void testSearchConstants() {
        assertEquals("confirmReenterSearch", SubCatalogConstants.CONFIRM_REENTER_SEARCH);
        assertEquals("handleReenterSearch", SubCatalogConstants.HANDLE_REENTER_SEARCH);
    }

    @Test
    void testCellDataConstants() {
        assertEquals("confirmReenterCellData", SubCatalogConstants.CONFIRM_REENTER_CELL_DATA);
        assertEquals("handleReenterCellData", SubCatalogConstants.HANDLE_REENTER_CELL_DATA);
    }

    @Test
    void testGetSharedExcelFileConstants() {
        assertEquals("confirmReenterGetSharedExcelFile", SubCatalogConstants.CONFIRM_REENTER_GET_SHARED_EXCEL_FILE);
        assertEquals("handleReenterGetSharedExcelFile", SubCatalogConstants.HANDLE_REENTER_GET_SHARED_EXCEL_FILE);
    }

    @Test
    void testGetColumnValuesConstants() {
        assertEquals("confirmReenterGetColumnValues", SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES);
        assertEquals("handleReenterGetColumnValues", SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES);
    }

    @Test
    void testColumnsSumConstants() {
        assertEquals("confirmReenterColumnsSum", SubCatalogConstants.CONFIRM_REENTER_COLUMNS_SUM);
        assertEquals("handleReenterColumnsSum", SubCatalogConstants.HANDLE_REENTER_COLUMNS_SUM);
    }

    @Test
    void testConvertToPdfConstants() {
        assertEquals("confirmReenterConvertToPdf", SubCatalogConstants.CONFIRM_REENTER_CONVERT_TO_PDF);
        assertEquals("handleReenterConvertToPdf", SubCatalogConstants.HANDLE_REENTER_CONVERT_TO_PDF);
    }

    @Test
    void testDeleteColumnWithIndexConstants() {
        assertEquals("confirmReenterDeleteColumnWithIndex", SubCatalogConstants.CONFIRM_REENTER_DELETE_COLUMN_WITH_INDEX);
        assertEquals("handleReenterDeleteColumnWithIndex", SubCatalogConstants.HANDLE_REENTER_DELETE_COLUMN_WITH_INDEX);
        assertEquals("confirmReenterDeleteColumnWithIndex", SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX);
        assertEquals("handleReenterDeleteColumnWithIndex", SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX);
    }

    @Test
    void testAddRowWithIndexConstants() {
        assertEquals("confirmReenterAddRowWithIndex", SubCatalogConstants.CONFIRM_REENTER_ADD_ROW_WITH_INDEX);
        assertEquals("handleReenterAddRowWithIndex", SubCatalogConstants.HANDLE_REENTER_ADD_ROW_WITH_INDEX);
        assertEquals("confirmReenterAddRowWithIndex", SubCatalogConstants.CONFIRM_REENTER_ADD_SHEET_ROW_WITH_INDEX);
        assertEquals("handleReenterAddRowWithIndex", SubCatalogConstants.HANDLE_REENTER_ADD_SHEET_ROW_WITH_INDEX);
    }

    @Test
    void testUpdateRowWithIndexConstants() {
        assertEquals("confirmReenterUpdateRowWithIndex", SubCatalogConstants.CONFIRM_REENTER_UPDATE_ROW_WITH_INDEX);
        assertEquals("handleReenterUpdateRowWithIndex", SubCatalogConstants.HANDLE_REENTER_UPDATE_ROW_WITH_INDEX);
        assertEquals("confirmReenterUpdateRowWithIndex", SubCatalogConstants.CONFIRM_REENTER_UPDATE_SHEET_ROW_WITH_INDEX);
        assertEquals("handleReenterUpdateRowWithIndex", SubCatalogConstants.HANDLE_REENTER_UPDATE_SHEET_ROW_WITH_INDEX);
    }

    @Test
    void testGetColumnValuesInSheetConstants() {
        assertEquals("confirmReenterGetColumnValuesInSheet", SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES_IN_SHEET);
        assertEquals("handleReenterGetColumnValuesInSheet", SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES_IN_SHEET);
        assertEquals("confirmReenterGetColumnValuesInSheet", SubCatalogConstants.CONFIRM_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET);
        assertEquals("handleReenterGetColumnValuesInSheet", SubCatalogConstants.HANDLE_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET);
    }

    @Test
    void testAddMultipleCellDataConstants() {
        assertEquals("confirmReenterAddMultipleCellData", SubCatalogConstants.CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA);
        assertEquals("handleReenterAddMultipleCellData", SubCatalogConstants.HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA);
        assertEquals("confirmReenterAddMultipleCellData", SubCatalogConstants.CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET);
        assertEquals("handleReenterAddMultipleCellData", SubCatalogConstants.HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET);
    }

    @Test
    void testClearSheetConstants() {
        assertEquals("confirmReenterClearSharedSheet", SubCatalogConstants.CONFIRM_REENTER_CLEAR_SHARED_SHEET);
        assertEquals("handleReenterClearSharedSheet", SubCatalogConstants.HANDLE_REENTER_CLEAR_SHARED_SHEET);
        assertEquals("confirmReenterClearSharedSheet", SubCatalogConstants.CONFIRM_REENTER_CLEAR_SHEET);
        assertEquals("handleReenterClearSharedSheet", SubCatalogConstants.HANDLE_REENTER_CLEAR_SHEET);
    }

    @Test
    void testDeleteSheetConstants() {
        assertEquals("confirmReenterDeleteSheet", SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET);
        assertEquals("handleReenterDeleteSheet", SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET);
    }

    @Test
    void testGetCellValueConstants() {
        assertEquals("confirmReenterGetCellValue", SubCatalogConstants.CONFIRM_REENTER_GET_CELL_VALUE);
        assertEquals("handleReenterGetCellValue", SubCatalogConstants.HANDLE_REENTER_GET_CELL_VALUE);
    }

    @Test
    void testAddRowsConstants() {
        assertEquals("confirmReenterAddRows", SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS);
        assertEquals("handleReenterAddRows", SubCatalogConstants.HANDLE_REENTER_ADD_ROWS);
    }

    @Test
    void testFindValueConstants() {
        assertEquals("confirmReenterFindValue", SubCatalogConstants.CONFIRM_REENTER_FIND_VALUE);
        assertEquals("handleReenterFindValue", SubCatalogConstants.HANDLE_REENTER_FIND_VALUE);
    }

    @Test
    void testGetExcelFileInfoConstants() {
        assertEquals("confirmReenterGetExcelFileInfo", SubCatalogConstants.CONFIRM_REENTER_GET_EXCEL_FILE_INFO);
        assertEquals("handleReenterGetExcelFileInfo", SubCatalogConstants.HANDLE_REENTER_GET_EXCEL_FILE_INFO);
    }

    @Test
    void testGetSheetsInfoConstants() {
        assertEquals("confirmReenterGetSheetsInfo", SubCatalogConstants.CONFIRM_REENTER_GET_SHEETS_INFO);
        assertEquals("handleReenterGetSheetsInfo", SubCatalogConstants.HANDLE_REENTER_GET_SHEETS_INFO);
    }

    @Test
    void testRetryAddDataToCellConstants() {
        assertEquals("confirmRetryAddDataToCell", SubCatalogConstants.CONFIRM_RETRY_ADD_DATA_TO_CELL);
        assertEquals("handleRetryAddDataToCell", SubCatalogConstants.HANDLE_RETRY_ADD_DATA_TO_CELL);
    }

    @Test
    void testRetryExecuteFormulaConstants() {
        assertEquals("confirmRetryExecuteFormula", SubCatalogConstants.CONFIRM_RETRY_EXECUTE_FORMULA);
        assertEquals("handleRetryExecuteFormula", SubCatalogConstants.HANDLE_RETRY_EXECUTE_FORMULA);
    }

    @Test
    void testCreateSheetConstants() {
        assertEquals("confirmReenterCreateSheet", SubCatalogConstants.CONFIRM_REENTER_CREATE_SHEET);
        assertEquals("handleReenterCreateSheet", SubCatalogConstants.HANDLE_REENTER_CREATE_SHEET);
    }

    @Test
    void testAddFormulaToCellConstants() {
        assertEquals("confirmReenterAddFormulaToCell", SubCatalogConstants.CONFIRM_REENTER_ADD_FORMULA_TO_CELL);
        assertEquals("handleReenterAddFormulaToCell", SubCatalogConstants.HANDLE_REENTER_ADD_FORMULA_TO_CELL);
    }

    @Test
    void testSummationConstants() {
        assertEquals("confirmReenterSummation", SubCatalogConstants.CONFIRM_REENTER_SUMMATION);
        assertEquals("handleReenterSummation", SubCatalogConstants.HANDLE_REENTER_SUMMATION);
    }

    @Test
    void testUploadFileAndConvertToPdfConstants() {
        assertEquals("confirmReenterUploadFileAndConvertToPDF", SubCatalogConstants.CONFIRM_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF);
        assertEquals("handleReenterUploadFileAndConvertToPDF", SubCatalogConstants.HANDLE_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF);
    }

    @Test
    void testGetMinMaxValueConstants() {
        assertEquals("confirmReenterGetMinMaxValue", SubCatalogConstants.CONFIRM_REENTER_GET_MIN_MAX_VALUE);
        assertEquals("handleReenterGetMinMaxValue", SubCatalogConstants.HANDLE_REENTER_GET_MIN_MAX_VALUE);
    }

    @Test
    void testTableVlookupConstants() {
        assertEquals("confirmReenterTableVlookup", SubCatalogConstants.CONFIRM_REENTER_TABLE_VLOOKUP);
        assertEquals("handleReenterTableVlookup", SubCatalogConstants.HANDLE_REENTER_TABLE_VLOOKUP);
    }

    @Test
    void testAddRowsInTableConstants() {
        assertEquals("confirmReenterAddRowsInTable", SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_IN_TABLE);
        assertEquals("handleReenterAddRowsInTable", SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_IN_TABLE);
    }

    @Test
    void testAddRowsToTableWithIndexConstants() {
        assertEquals("confirmReenterAddRowsToTableWithIndex", SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX);
        assertEquals("handleReenterAddRowsToTableWithIndex", SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX);
    }

    @Test
    void testGetRowsConstants() {
        assertEquals("confirmReenterGetRows", SubCatalogConstants.CONFIRM_REENTER_GET_ROWS);
        assertEquals("handleReenterGetRows", SubCatalogConstants.HANDLE_REENTER_GET_ROWS);
    }

    @Test
    void testTableSummationConstants() {
        assertEquals("confirmReenterTableSummation", SubCatalogConstants.CONFIRM_REENTER_TABLE_SUMMATION);
        assertEquals("handleReenterTableSummation", SubCatalogConstants.HANDLE_REENTER_TABLE_SUMMATION);
    }

    @Test
    void testDeleteTableColumnConstants() {
        assertEquals("confirmReenterDeleteTableColumn", SubCatalogConstants.CONFIRM_REENTER_DELETE_TABLE_COLUMN);
        assertEquals("handleReenterDeleteTableColumn", SubCatalogConstants.HANDLE_REENTER_DELETE_TABLE_COLUMN);
    }

    @Test
    void testCreateTableForRangeConstants() {
        assertEquals("confirmReenterCreateTableForRange", SubCatalogConstants.CONFIRM_REENTER_CREATE_TABLE_FOR_RANGE);
        assertEquals("handleReenterCreateTableForRange", SubCatalogConstants.HANDLE_REENTER_CREATE_TABLE_FOR_RANGE);
    }

    @Test
    void testGetColumnValuesForKeysConstants() {
        assertEquals("confirmReenterGetColumnValuesForKeys", SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES_FOR_KEYS);
        assertEquals("handleReenterGetColumnValuesForKeys", SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES_FOR_KEYS);
    }

    @Test
    void testReenterWasTrueHenceContinuing() {
        assertEquals("Reenter was true, hence continuing", SubCatalogConstants.REENTER_WAS_TRUE_HENCE_CONTINUING);
    }

    @Test
    void testPrivateConstructor() throws Exception {
        Constructor<SubCatalogConstants> constructor = SubCatalogConstants.class.getDeclaredConstructor();
        assertTrue(java.lang.reflect.Modifier.isPrivate(constructor.getModifiers()));
        constructor.setAccessible(true);
        assertNotNull(constructor.newInstance());
    }

    @Test
    void testAllConstantsAreNonNull() {
        assertNotNull(SubCatalogConstants.VALIDATION_RESULTS);
        assertNotNull(SubCatalogConstants.REENTER);
        assertNotNull(SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_TABLE);
        assertNotNull(SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_TABLE);
        assertNotNull(SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_SHEET);
        assertNotNull(SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_SHEET);
        assertNotNull(SubCatalogConstants.CONFIRM_REENTER_RANGE_SUM);
        assertNotNull(SubCatalogConstants.HANDLE_REENTER_RANGE_SUM);
        assertNotNull(SubCatalogConstants.REENTER_WAS_TRUE_HENCE_CONTINUING);
    }
}
