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

/**
 * Constants for sub catalog request names and common fields used in Office365Excel extension.
 */
public class SubCatalogConstants {
    
    private SubCatalogConstants() {}
    
    // Validation constants
    public static final String VALIDATION_RESULTS = "VALIDATION_RESULTS";
    public static final String REENTER = "REENTER";
    
    // Sub catalog request names
    public static final String CONFIRM_REENTER_ADD_ROWS_TO_TABLE = "confirmReenterAddRowsToTable";
    public static final String HANDLE_REENTER_ADD_ROWS_TO_TABLE = "handleReenterAddRowsToTable";
    
    public static final String CONFIRM_REENTER_ADD_ROWS_TO_SHEET = "confirmReenterAddRowsToSheet";
    public static final String HANDLE_REENTER_ADD_ROWS_TO_SHEET = "handleReenterAddRowsToSheet";

    public static final String CONFIRM_REENTER_RANGE_SUM = "confirmReenterRangeSum";
    public static final String HANDLE_REENTER_RANGE_SUM = "handleReenterRangeSum";

    public static final String CONFIRM_REENTER_GET_TABLE_ROWS = "confirmReenterGetTableRows";
    public static final String HANDLE_REENTER_GET_TABLE_ROWS = "handleReenterGetTableRows";

    public static final String CONFIRM_REENTER_GET_ROWS_FOR_RANGE = "confirmReenterGetRowsForRange";
    public static final String HANDLE_REENTER_GET_ROWS_FOR_RANGE = "handleReenterGetRowsForRange";

    public static final String CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET = "confirmReenterGetRowsForRangeForSheet";
    public static final String HANDLE_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET = "handleReenterGetRowsForRangeForSheet";

    public static final String CONFIRM_REENTER_SHARED_SHEET_VLOOKUP = "confirmReenterSheetVlookup";
    public static final String HANDLE_REENTER_SHARED_SHEET_VLOOKUP = "handleReenterSheetVlookup";

    public static final String CONFIRM_REENTER_SHEET_VLOOKUP = "confirmReenterSheetVlookup";
    public static final String HANDLE_REENTER_SHEET_VLOOKUP = "handleReenterSheetVlookup";

    public static final String CONFIRM_REENTER_ADD_DATA_TO_CELL = "confirmReenterAddDataToCell";
    public static final String HANDLE_REENTER_ADD_DATA_TO_CELL = "handleReenterAddDataToCell";

    public static final String CONFIRM_REENTER_VLOOKUP_IN_TABLE = "confirmReenterVlookupInTable";
    public static final String HANDLE_REENTER_VLOOKUP_IN_TABLE = "handleReenterVlookupInTable";

    public static final String CONFIRM_REENTER_DELETE_ROW_WITH_INDEX = "confirmReenterDeleteRowWithIndex";
    public static final String HANDLE_REENTER_DELETE_ROW_WITH_INDEX = "handleReenterDeleteRowWithIndex";

    public static final String CONFIRM_REENTER_DELETE_SHEET_ROW_WITH_INDEX = "confirmReenterDeleteRowWithIndex";
    public static final String HANDLE_REENTER_DELETE_SHEET_ROW_WITH_INDEX = "handleReenterDeleteRowWithIndex";

    public static final String CONFIRM_REENTER_GET_COMPLETE_ROW_INFO = "confirmReenterGetCompleteRowInfo";
    public static final String HANDLE_REENTER_GET_COMPLETE_ROW_INFO = "handleReenterGetCompleteRowInfo";

    public static final String CONFIRM_REENTER_GET_COMPLETE_SHEET_ROW_INFO = "confirmReenterGetCompleteRowInfo";
    public static final String HANDLE_REENTER_GET_COMPLETE_SHEET_ROW_INFO = "handleReenterGetCompleteRowInfo";

    public static final String CONFIRM_REENTER_DELETE_COLUMN = "confirmReenterDeleteColumn";
    public static final String HANDLE_REENTER_DELETE_COLUMN = "handleReenterDeleteColumn";

    public static final String CONFIRM_REENTER_EXECUTE_FORMULA = "confirmReenterExecuteFormula";
    public static final String HANDLE_REENTER_EXECUTE_FORMULA = "handleReenterExecuteFormula";

    public static final String CONFIRM_REENTER_SEARCH = "confirmReenterSearch";
    public static final String HANDLE_REENTER_SEARCH = "handleReenterSearch";

    public static final String CONFIRM_REENTER_CELL_DATA = "confirmReenterCellData";
    public static final String HANDLE_REENTER_CELL_DATA = "handleReenterCellData";

    public static final String CONFIRM_REENTER_GET_SHARED_EXCEL_FILE = "confirmReenterGetSharedExcelFile";
    public static final String HANDLE_REENTER_GET_SHARED_EXCEL_FILE = "handleReenterGetSharedExcelFile";

    public static final String CONFIRM_REENTER_GET_COLUMN_VALUES = "confirmReenterGetColumnValues";
    public static final String HANDLE_REENTER_GET_COLUMN_VALUES = "handleReenterGetColumnValues";

    public static final String CONFIRM_REENTER_COLUMNS_SUM = "confirmReenterColumnsSum";
    public static final String HANDLE_REENTER_COLUMNS_SUM = "handleReenterColumnsSum";

    public static final String CONFIRM_REENTER_CONVERT_TO_PDF = "confirmReenterConvertToPdf";
    public static final String HANDLE_REENTER_CONVERT_TO_PDF = "handleReenterConvertToPdf";

    public static final String CONFIRM_REENTER_DELETE_COLUMN_WITH_INDEX = "confirmReenterDeleteColumnWithIndex";
    public static final String HANDLE_REENTER_DELETE_COLUMN_WITH_INDEX = "handleReenterDeleteColumnWithIndex";

    public static final String CONFIRM_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX = "confirmReenterDeleteColumnWithIndex";
    public static final String HANDLE_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX = "handleReenterDeleteColumnWithIndex";

    public static final String CONFIRM_REENTER_ADD_ROW_WITH_INDEX = "confirmReenterAddRowWithIndex";
    public static final String HANDLE_REENTER_ADD_ROW_WITH_INDEX = "handleReenterAddRowWithIndex";

    public static final String CONFIRM_REENTER_ADD_SHEET_ROW_WITH_INDEX = "confirmReenterAddRowWithIndex";
    public static final String HANDLE_REENTER_ADD_SHEET_ROW_WITH_INDEX = "handleReenterAddRowWithIndex";

    public static final String CONFIRM_REENTER_UPDATE_ROW_WITH_INDEX = "confirmReenterUpdateRowWithIndex";
    public static final String HANDLE_REENTER_UPDATE_ROW_WITH_INDEX = "handleReenterUpdateRowWithIndex";

    public static final String CONFIRM_REENTER_UPDATE_SHEET_ROW_WITH_INDEX = "confirmReenterUpdateRowWithIndex";
    public static final String HANDLE_REENTER_UPDATE_SHEET_ROW_WITH_INDEX = "handleReenterUpdateRowWithIndex";

    public static final String CONFIRM_REENTER_GET_COLUMN_VALUES_IN_SHEET = "confirmReenterGetColumnValuesInSheet";
    public static final String HANDLE_REENTER_GET_COLUMN_VALUES_IN_SHEET = "handleReenterGetColumnValuesInSheet";

    public static final String CONFIRM_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET = "confirmReenterGetColumnValuesInSheet";
    public static final String HANDLE_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET = "handleReenterGetColumnValuesInSheet";

    public static final String CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA = "confirmReenterAddMultipleCellData";
    public static final String HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA = "handleReenterAddMultipleCellData";

    public static final String CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET = "confirmReenterAddMultipleCellData";
    public static final String HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET = "handleReenterAddMultipleCellData";

    public static final String CONFIRM_REENTER_CLEAR_SHARED_SHEET = "confirmReenterClearSharedSheet";
    public static final String HANDLE_REENTER_CLEAR_SHARED_SHEET = "handleReenterClearSharedSheet";

    public static final String CONFIRM_REENTER_CLEAR_SHEET = "confirmReenterClearSharedSheet";
    public static final String HANDLE_REENTER_CLEAR_SHEET = "handleReenterClearSharedSheet";

    public static final String CONFIRM_REENTER_DELETE_SHEET = "confirmReenterDeleteSheet";
    public static final String HANDLE_REENTER_DELETE_SHEET = "handleReenterDeleteSheet";

    public static final String CONFIRM_REENTER_GET_CELL_VALUE = "confirmReenterGetCellValue";
    public static final String HANDLE_REENTER_GET_CELL_VALUE = "handleReenterGetCellValue";

    public static final String CONFIRM_REENTER_ADD_ROWS = "confirmReenterAddRows";
    public static final String HANDLE_REENTER_ADD_ROWS = "handleReenterAddRows";

    public static final String CONFIRM_REENTER_FIND_VALUE = "confirmReenterFindValue";
    public static final String HANDLE_REENTER_FIND_VALUE = "handleReenterFindValue";

    public static final String CONFIRM_REENTER_GET_EXCEL_FILE_INFO = "confirmReenterGetExcelFileInfo";
    public static final String HANDLE_REENTER_GET_EXCEL_FILE_INFO = "handleReenterGetExcelFileInfo";

    public static final String CONFIRM_REENTER_GET_SHEETS_INFO = "confirmReenterGetSheetsInfo";
    public static final String HANDLE_REENTER_GET_SHEETS_INFO = "handleReenterGetSheetsInfo";

    public static final String CONFIRM_RETRY_ADD_DATA_TO_CELL = "confirmRetryAddDataToCell";
    public static final String HANDLE_RETRY_ADD_DATA_TO_CELL = "handleRetryAddDataToCell";

    public static final String CONFIRM_RETRY_EXECUTE_FORMULA = "confirmRetryExecuteFormula";
    public static final String HANDLE_RETRY_EXECUTE_FORMULA = "handleRetryExecuteFormula";

    public static final String CONFIRM_REENTER_CREATE_SHEET = "confirmReenterCreateSheet";
    public static final String HANDLE_REENTER_CREATE_SHEET = "handleReenterCreateSheet";

    public static final String CONFIRM_REENTER_ADD_FORMULA_TO_CELL = "confirmReenterAddFormulaToCell";
    public static final String HANDLE_REENTER_ADD_FORMULA_TO_CELL = "handleReenterAddFormulaToCell";

    public static final String CONFIRM_REENTER_SUMMATION = "confirmReenterSummation";
    public static final String HANDLE_REENTER_SUMMATION = "handleReenterSummation";

    public static final String CONFIRM_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF = "confirmReenterUploadFileAndConvertToPDF";
    public static final String HANDLE_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF = "handleReenterUploadFileAndConvertToPDF";

    public static final String CONFIRM_REENTER_GET_MIN_MAX_VALUE = "confirmReenterGetMinMaxValue";
    public static final String HANDLE_REENTER_GET_MIN_MAX_VALUE = "handleReenterGetMinMaxValue";

    public static final String CONFIRM_REENTER_TABLE_VLOOKUP = "confirmReenterTableVlookup";
    public static final String HANDLE_REENTER_TABLE_VLOOKUP = "handleReenterTableVlookup";

    public static final String CONFIRM_REENTER_ADD_ROWS_IN_TABLE = "confirmReenterAddRowsInTable";
    public static final String HANDLE_REENTER_ADD_ROWS_IN_TABLE = "handleReenterAddRowsInTable";

    public static final String CONFIRM_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX = "confirmReenterAddRowsToTableWithIndex";
    public static final String HANDLE_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX = "handleReenterAddRowsToTableWithIndex";

    public static final String CONFIRM_REENTER_GET_ROWS = "confirmReenterGetRows";
    public static final String HANDLE_REENTER_GET_ROWS = "handleReenterGetRows";

    public static final String CONFIRM_REENTER_TABLE_SUMMATION = "confirmReenterTableSummation";
    public static final String HANDLE_REENTER_TABLE_SUMMATION = "handleReenterTableSummation";

    public static final String CONFIRM_REENTER_DELETE_TABLE_COLUMN = "confirmReenterDeleteTableColumn";
    public static final String HANDLE_REENTER_DELETE_TABLE_COLUMN = "handleReenterDeleteTableColumn";

    public static final String CONFIRM_REENTER_CREATE_TABLE_FOR_RANGE = "confirmReenterCreateTableForRange";
    public static final String HANDLE_REENTER_CREATE_TABLE_FOR_RANGE = "handleReenterCreateTableForRange";

    public static final String CONFIRM_REENTER_GET_COLUMN_VALUES_FOR_KEYS = "confirmReenterGetColumnValuesForKeys";
    public static final String HANDLE_REENTER_GET_COLUMN_VALUES_FOR_KEYS = "handleReenterGetColumnValuesForKeys";

    // Step messages
    public static final String REENTER_WAS_TRUE_HENCE_CONTINUING = "Reenter was true, hence continuing";
}
