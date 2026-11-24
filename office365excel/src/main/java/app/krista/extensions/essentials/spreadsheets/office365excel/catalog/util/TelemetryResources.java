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
 * Constants for telemetry operation types and tag keys used in Office365Excel extension.
 * This centralizes all telemetry string literals to prevent duplication and typos.
 */
public class TelemetryResources {

    // Operation Type Constants
    public static final String TABLE_INSERT = "table_insert";
    public static final String SHEET_INSERT = "sheet_insert";
    public static final String RANGE_SUM = "range_sum";
    public static final String VLOOKUP = "vlookup";
    public static final String TABLE_QUERY = "table_query";
    public static final String CELL_UPDATE = "cell_update";
    public static final String DELETE_COLUMN = "delete_column";
    public static final String COLUMN_VALUES_FOR_KEYS = "column_values_for_keys";
    public static final String EXECUTE_FORMULA = "execute_formula";
    public static final String VLOOKUP_IN_TABLE = "vlookup_in_table";
    public static final String COLUMNS_SUM = "columns_sum";
    public static final String SEARCH = "search";
    public static final String ADD_ROWS_IN_TABLE = "add_rows_in_table";
    public static final String DELETE_TABLE_COLUMN = "delete_table_column";
    public static final String CELL_DATA = "cell_data";
    public static final String ADD_ROW_WITH_INDEX = "add_row_with_index";
    public static final String UPDATE_ROW_WITH_INDEX = "update_row_with_index";
    public static final String GET_COMPLETE_ROW_INFO = "get_complete_row_info";
    public static final String ADD_MULTIPLE_CELL_DATA = "add_multiple_cell_data";
    public static final String GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET = "get_column_values_for_keys_in_sheet";
    public static final String GET_CELL_VALUE = "get_cell_value";
    public static final String ADD_ROWS = "add_rows";
    public static final String FIND_VALUE = "find_value";
    public static final String GET_EXCEL_FILE_INFO = "get_excel_file_info";
    public static final String GET_SHEETS_INFO = "get_sheets_info";
    public static final String ADD_DATA_TO_CELL = "add_data_to_cell";
    public static final String CREATE_SHEET = "create_sheet";
    public static final String CLEAR_SHEET_FOR_RANGE = "clear_sheet_for_range";
    public static final String ADD_FORMULA_TO_CELL = "add_formula_to_cell";
    public static final String GET_MIN_MAX_VALUE = "get_min_max_value";
    public static final String TABLE_VLOOKUP = "table_vlookup";
    public static final String TABLE_CREATION = "table_creation";
    public static final String ADD_ROWS_TO_TABLE_WITH_INDEX = "add_rows_to_table_with_index";
    public static final String TABLE_SUMMATION = "table_summation";
    public static final String DELETE_COLUMN_WITH_INDEX = "delete_column_with_index";
    public static final String DELETE_ROW_WITH_INDEX = "delete_row_with_index";
    public static final String GET_SHARED_EXCEL_FILE = "get_shared_excel_file";
    public static final String CONVERT_TO_PDF = "convert_to_pdf";
    public static final String CLEAR_SHARED_SHEET_FOR_RANGE = "clear_shared_sheet_for_range";
    public static final String DELETE_SHEET = "delete_sheet";
    public static final String UPLOAD_FILE_AND_CONVERT_TO_PDF = "upload_file_and_convert_to_pdf";
    public static final String GET_COLUMN_VALUES_FOR_KEYS_OPERATION = "get_column_values_for_keys_operation";

    // Confirmation Operation Type Constants
    public static final String ADD_DATA_TO_CELL_CONFIRMATION = "add_data_to_cell_confirmation";
    public static final String CELL_DATA_CONFIRMATION = "cell_data_confirmation";
    public static final String ADD_MULTIPLE_CELL_DATA_CONFIRMATION = "add_multiple_cell_data_confirmation";
    public static final String GET_CELL_VALUE_CONFIRMATION = "get_cell_value_confirmation";
    public static final String ADD_DATA_TO_CELL_RETRY_CONFIRMATION = "add_data_to_cell_retry_confirmation";
    public static final String GET_COLUMN_VALUES_CONFIRMATION = "get_column_values_confirmation";
    public static final String COLUMNS_SUM_CONFIRMATION = "columns_sum_confirmation";
    public static final String DELETE_COLUMN_WITH_INDEX_CONFIRMATION = "delete_column_with_index_confirmation";
    public static final String GET_COLUMN_VALUES_FOR_KEYS_CONFIRMATION = "get_column_values_for_keys_confirmation";
    public static final String GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET_CONFIRMATION = "get_column_values_for_keys_in_sheet_confirmation";
    public static final String UPLOAD_FILE_AND_CONVERT_TO_PDF_CONFIRMATION = "upload_file_and_convert_to_pdf_confirmation";
    public static final String UPLOAD_FILE_AND_CONVERT_TO_PDF_OPERATION = "upload_file_and_convert_to_pdf_operation";

    // Tag Key Constants
    public static final String OPERATION_TYPE = "operationType";
    public static final String HAS_ROWS = "hasRows";
    public static final String HAS_CSV_FILE = "hasCsvFile";
    public static final String ROW_COUNT = "rowCount";
    public static final String CSV_FILE_COUNT = "csvFileCount";
    public static final String HAS_FILE = "hasFile";
    public static final String FILE_COUNT = "fileCount";
    public static final String HAS_RANGE = "hasRange";
    public static final String SHEET_PROVIDED = "sheetProvided";
    public static final String HAS_LOOKUP_VALUE = "hasLookupValue";
    public static final String HAS_TABLE_NAME = "hasTableName";
    public static final String HAS_CELL_ID = "hasCellId";
    public static final String HAS_DATA = "hasData";
    public static final String DATA_LENGTH = "dataLength";
    public static final String HAS_COLUMN_NAME = "hasColumnName";
    public static final String KEY_ROWS_COUNT = "keyRowsCount";
    public static final String LOOKUP_COLUMN_COUNT = "lookupColumnCount";
    public static final String HAS_FORMULA = "hasFormula";
    public static final String HAS_FORMULA_NAME = "hasFormulaName";
    public static final String FORMULA_LENGTH = "formulaLength";
    public static final String HAS_COLUMN_INDEX = "hasColumnIndex";
    public static final String COLUMN_COUNT = "columnCount";
    public static final String HAS_COLUMNS = "hasColumns";
    public static final String HAS_SHEET_NAME = "hasSheetName";
    public static final String HAS_DATA_TO_SEARCH = "hasDataToSearch";
    public static final String HAS_ROW_INDEX = "hasRowIndex";
    public static final String HAS_ROW_DATA = "hasRowData";
    public static final String HAS_CELL_DATA = "hasCellData";
    public static final String CELL_DATA_COUNT = "cellDataCount";
    public static final String HAS_LOOKUP_ROWS = "hasLookupRows";
    public static final String HAS_LOOKUP_COLUMNS = "hasLookupColumns";
    public static final String LOOKUP_ROWS_COUNT = "lookupRowsCount";
    public static final String LOOKUP_COLUMNS_COUNT = "lookupColumnsCount";
    public static final String GET_ALL_RECORDS = "getAllRecords";
    public static final String HAS_COLUMN_NUMBER = "hasColumnNumber";
    public static final String HAS_FILE_ID = "hasFileId";
    public static final String HAS_PATH = "hasPath";
    public static final String HAS_NAME = "hasName";
    public static final String HAS_CELL_ID_UPPER = "hasCellID";
    public static final String HAS_CELL_RANGE = "hasCellRange";
    public static final String HAS_KEY_COLUMN_NAME = "hasKeyColumnName";
    public static final String HAS_COLUMN_RANGE = "hasColumnRange";
    public static final String HAS_INDEX = "hasIndex";
    public static final String HAS_SHARED_FOLDER_NAME = "hasSharedFolderName";
    public static final String HAS_SHARED_FILE_NAME = "hasSharedFileName";
    public static final String HAS_DRIVE_ID = "hasDriveId";
    public static final String HAS_EXCEL_FILE = "hasExcelFile";
    public static final String HAS_FILE_PATH = "hasFilePath";

}