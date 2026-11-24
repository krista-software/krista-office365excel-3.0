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

import static org.junit.jupiter.api.Assertions.*;

class TelemetryResourcesTest {

    // Operation Type Constants
    @Test
    void testTableInsert() {
        assertEquals("table_insert", TelemetryResources.TABLE_INSERT);
    }

    @Test
    void testSheetInsert() {
        assertEquals("sheet_insert", TelemetryResources.SHEET_INSERT);
    }

    @Test
    void testRangeSum() {
        assertEquals("range_sum", TelemetryResources.RANGE_SUM);
    }

    @Test
    void testVlookup() {
        assertEquals("vlookup", TelemetryResources.VLOOKUP);
    }

    @Test
    void testTableQuery() {
        assertEquals("table_query", TelemetryResources.TABLE_QUERY);
    }

    @Test
    void testCellUpdate() {
        assertEquals("cell_update", TelemetryResources.CELL_UPDATE);
    }

    @Test
    void testDeleteColumn() {
        assertEquals("delete_column", TelemetryResources.DELETE_COLUMN);
    }

    @Test
    void testColumnValuesForKeys() {
        assertEquals("column_values_for_keys", TelemetryResources.COLUMN_VALUES_FOR_KEYS);
    }

    @Test
    void testExecuteFormula() {
        assertEquals("execute_formula", TelemetryResources.EXECUTE_FORMULA);
    }

    @Test
    void testVlookupInTable() {
        assertEquals("vlookup_in_table", TelemetryResources.VLOOKUP_IN_TABLE);
    }

    @Test
    void testColumnsSum() {
        assertEquals("columns_sum", TelemetryResources.COLUMNS_SUM);
    }

    @Test
    void testSearch() {
        assertEquals("search", TelemetryResources.SEARCH);
    }

    @Test
    void testAddRowsInTable() {
        assertEquals("add_rows_in_table", TelemetryResources.ADD_ROWS_IN_TABLE);
    }

    @Test
    void testDeleteTableColumn() {
        assertEquals("delete_table_column", TelemetryResources.DELETE_TABLE_COLUMN);
    }

    @Test
    void testCellData() {
        assertEquals("cell_data", TelemetryResources.CELL_DATA);
    }

    @Test
    void testAddRowWithIndex() {
        assertEquals("add_row_with_index", TelemetryResources.ADD_ROW_WITH_INDEX);
    }

    @Test
    void testUpdateRowWithIndex() {
        assertEquals("update_row_with_index", TelemetryResources.UPDATE_ROW_WITH_INDEX);
    }

    @Test
    void testGetCompleteRowInfo() {
        assertEquals("get_complete_row_info", TelemetryResources.GET_COMPLETE_ROW_INFO);
    }

    @Test
    void testAddMultipleCellData() {
        assertEquals("add_multiple_cell_data", TelemetryResources.ADD_MULTIPLE_CELL_DATA);
    }

    @Test
    void testGetColumnValuesForKeysInSheet() {
        assertEquals("get_column_values_for_keys_in_sheet", TelemetryResources.GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET);
    }

    @Test
    void testGetCellValue() {
        assertEquals("get_cell_value", TelemetryResources.GET_CELL_VALUE);
    }

    @Test
    void testAddRows() {
        assertEquals("add_rows", TelemetryResources.ADD_ROWS);
    }

    @Test
    void testFindValue() {
        assertEquals("find_value", TelemetryResources.FIND_VALUE);
    }

    @Test
    void testGetExcelFileInfo() {
        assertEquals("get_excel_file_info", TelemetryResources.GET_EXCEL_FILE_INFO);
    }

    @Test
    void testGetSheetsInfo() {
        assertEquals("get_sheets_info", TelemetryResources.GET_SHEETS_INFO);
    }

    @Test
    void testAddDataToCell() {
        assertEquals("add_data_to_cell", TelemetryResources.ADD_DATA_TO_CELL);
    }

    @Test
    void testCreateSheet() {
        assertEquals("create_sheet", TelemetryResources.CREATE_SHEET);
    }

    @Test
    void testClearSheetForRange() {
        assertEquals("clear_sheet_for_range", TelemetryResources.CLEAR_SHEET_FOR_RANGE);
    }

    @Test
    void testAddFormulaToCell() {
        assertEquals("add_formula_to_cell", TelemetryResources.ADD_FORMULA_TO_CELL);
    }

    @Test
    void testGetMinMaxValue() {
        assertEquals("get_min_max_value", TelemetryResources.GET_MIN_MAX_VALUE);
    }

    @Test
    void testTableVlookup() {
        assertEquals("table_vlookup", TelemetryResources.TABLE_VLOOKUP);
    }

    @Test
    void testTableCreation() {
        assertEquals("table_creation", TelemetryResources.TABLE_CREATION);
    }

    @Test
    void testAddRowsToTableWithIndex() {
        assertEquals("add_rows_to_table_with_index", TelemetryResources.ADD_ROWS_TO_TABLE_WITH_INDEX);
    }

    @Test
    void testTableSummation() {
        assertEquals("table_summation", TelemetryResources.TABLE_SUMMATION);
    }

    @Test
    void testDeleteColumnWithIndex() {
        assertEquals("delete_column_with_index", TelemetryResources.DELETE_COLUMN_WITH_INDEX);
    }

    @Test
    void testDeleteRowWithIndex() {
        assertEquals("delete_row_with_index", TelemetryResources.DELETE_ROW_WITH_INDEX);
    }

    @Test
    void testGetSharedExcelFile() {
        assertEquals("get_shared_excel_file", TelemetryResources.GET_SHARED_EXCEL_FILE);
    }

    @Test
    void testConvertToPdf() {
        assertEquals("convert_to_pdf", TelemetryResources.CONVERT_TO_PDF);
    }

    @Test
    void testClearSharedSheetForRange() {
        assertEquals("clear_shared_sheet_for_range", TelemetryResources.CLEAR_SHARED_SHEET_FOR_RANGE);
    }

    @Test
    void testDeleteSheet() {
        assertEquals("delete_sheet", TelemetryResources.DELETE_SHEET);
    }

    @Test
    void testUploadFileAndConvertToPdf() {
        assertEquals("upload_file_and_convert_to_pdf", TelemetryResources.UPLOAD_FILE_AND_CONVERT_TO_PDF);
    }

    @Test
    void testGetColumnValuesForKeysOperation() {
        assertEquals("get_column_values_for_keys_operation", TelemetryResources.GET_COLUMN_VALUES_FOR_KEYS_OPERATION);
    }

    // Confirmation Operation Type Constants
    @Test
    void testConfirmationConstants() {
        assertEquals("add_data_to_cell_confirmation", TelemetryResources.ADD_DATA_TO_CELL_CONFIRMATION);
        assertEquals("cell_data_confirmation", TelemetryResources.CELL_DATA_CONFIRMATION);
        assertEquals("add_multiple_cell_data_confirmation", TelemetryResources.ADD_MULTIPLE_CELL_DATA_CONFIRMATION);
        assertEquals("get_cell_value_confirmation", TelemetryResources.GET_CELL_VALUE_CONFIRMATION);
        assertEquals("add_data_to_cell_retry_confirmation", TelemetryResources.ADD_DATA_TO_CELL_RETRY_CONFIRMATION);
        assertEquals("get_column_values_confirmation", TelemetryResources.GET_COLUMN_VALUES_CONFIRMATION);
        assertEquals("columns_sum_confirmation", TelemetryResources.COLUMNS_SUM_CONFIRMATION);
        assertEquals("delete_column_with_index_confirmation", TelemetryResources.DELETE_COLUMN_WITH_INDEX_CONFIRMATION);
        assertEquals("get_column_values_for_keys_confirmation", TelemetryResources.GET_COLUMN_VALUES_FOR_KEYS_CONFIRMATION);
        assertEquals("get_column_values_for_keys_in_sheet_confirmation", TelemetryResources.GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET_CONFIRMATION);
        assertEquals("upload_file_and_convert_to_pdf_confirmation", TelemetryResources.UPLOAD_FILE_AND_CONVERT_TO_PDF_CONFIRMATION);
        assertEquals("upload_file_and_convert_to_pdf_operation", TelemetryResources.UPLOAD_FILE_AND_CONVERT_TO_PDF_OPERATION);
    }

    // Tag Key Constants
    @Test
    void testTagKeyConstants() {
        assertEquals("operationType", TelemetryResources.OPERATION_TYPE);
        assertEquals("hasRows", TelemetryResources.HAS_ROWS);
        assertEquals("hasCsvFile", TelemetryResources.HAS_CSV_FILE);
        assertEquals("rowCount", TelemetryResources.ROW_COUNT);
        assertEquals("csvFileCount", TelemetryResources.CSV_FILE_COUNT);
        assertEquals("hasFile", TelemetryResources.HAS_FILE);
        assertEquals("fileCount", TelemetryResources.FILE_COUNT);
        assertEquals("hasRange", TelemetryResources.HAS_RANGE);
        assertEquals("sheetProvided", TelemetryResources.SHEET_PROVIDED);
        assertEquals("hasLookupValue", TelemetryResources.HAS_LOOKUP_VALUE);
        assertEquals("hasTableName", TelemetryResources.HAS_TABLE_NAME);
        assertEquals("hasCellId", TelemetryResources.HAS_CELL_ID);
        assertEquals("hasData", TelemetryResources.HAS_DATA);
        assertEquals("dataLength", TelemetryResources.DATA_LENGTH);
        assertEquals("hasColumnName", TelemetryResources.HAS_COLUMN_NAME);
        assertEquals("keyRowsCount", TelemetryResources.KEY_ROWS_COUNT);
        assertEquals("lookupColumnCount", TelemetryResources.LOOKUP_COLUMN_COUNT);
        assertEquals("hasFormula", TelemetryResources.HAS_FORMULA);
        assertEquals("hasFormulaName", TelemetryResources.HAS_FORMULA_NAME);
        assertEquals("formulaLength", TelemetryResources.FORMULA_LENGTH);
        assertEquals("hasColumnIndex", TelemetryResources.HAS_COLUMN_INDEX);
        assertEquals("columnCount", TelemetryResources.COLUMN_COUNT);
        assertEquals("hasColumns", TelemetryResources.HAS_COLUMNS);
        assertEquals("hasSheetName", TelemetryResources.HAS_SHEET_NAME);
        assertEquals("hasDataToSearch", TelemetryResources.HAS_DATA_TO_SEARCH);
        assertEquals("hasRowIndex", TelemetryResources.HAS_ROW_INDEX);
        assertEquals("hasRowData", TelemetryResources.HAS_ROW_DATA);
        assertEquals("hasCellData", TelemetryResources.HAS_CELL_DATA);
        assertEquals("cellDataCount", TelemetryResources.CELL_DATA_COUNT);
        assertEquals("hasLookupRows", TelemetryResources.HAS_LOOKUP_ROWS);
        assertEquals("hasLookupColumns", TelemetryResources.HAS_LOOKUP_COLUMNS);
        assertEquals("lookupRowsCount", TelemetryResources.LOOKUP_ROWS_COUNT);
        assertEquals("lookupColumnsCount", TelemetryResources.LOOKUP_COLUMNS_COUNT);
        assertEquals("getAllRecords", TelemetryResources.GET_ALL_RECORDS);
        assertEquals("hasColumnNumber", TelemetryResources.HAS_COLUMN_NUMBER);
        assertEquals("hasFileId", TelemetryResources.HAS_FILE_ID);
        assertEquals("hasPath", TelemetryResources.HAS_PATH);
        assertEquals("hasName", TelemetryResources.HAS_NAME);
        assertEquals("hasCellID", TelemetryResources.HAS_CELL_ID_UPPER);
        assertEquals("hasCellRange", TelemetryResources.HAS_CELL_RANGE);
        assertEquals("hasKeyColumnName", TelemetryResources.HAS_KEY_COLUMN_NAME);
        assertEquals("hasColumnRange", TelemetryResources.HAS_COLUMN_RANGE);
        assertEquals("hasIndex", TelemetryResources.HAS_INDEX);
        assertEquals("hasSharedFolderName", TelemetryResources.HAS_SHARED_FOLDER_NAME);
        assertEquals("hasSharedFileName", TelemetryResources.HAS_SHARED_FILE_NAME);
        assertEquals("hasDriveId", TelemetryResources.HAS_DRIVE_ID);
        assertEquals("hasExcelFile", TelemetryResources.HAS_EXCEL_FILE);
        assertEquals("hasFilePath", TelemetryResources.HAS_FILE_PATH);
    }
}
