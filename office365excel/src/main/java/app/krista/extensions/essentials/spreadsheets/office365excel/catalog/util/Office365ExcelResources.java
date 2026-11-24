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
 * Constants for field names and input types used in SharedArea catalog requests.
 * This centralizes all string literals to prevent duplication and typos.
 */
public class Office365ExcelResources {

    public static final String DRIVE_ID = "Drive Id";
    public static final String FILE_ID = "File Id";
    public static final String TABLE_NAME = "Table Name";
    public static final String STATE_ID = "State Id";
    public static final String ROWS = "Rows";
    public static final String ROWS_FROM_CSV_FILE = "CSV File";
    public static final String RANGE = "Range";
    public static final String LOOKUP_VALUE = "Lookup Value";
    public static final String COLUMN_INDEX = "Column Index";
    public static final String EXACT_MATCH = "Exact Match";
    public static final String CELL_ID = "Cell Id";
    public static final String DATA = "Data";
    public static final String ROW_INDEX = "Row Index";
    public static final String COLUMN_NAME = "Column Name";
    public static final String FORMULA = "Formula";
    public static final String FORMULA_NAME = "Formula Name";
    public static final String COMMENT = "Comment";
    public static final String SHEET_NAME = "Sheet Name";
    public static final String DATA_TO_SEARCH = "Data to Search";
    public static final String SHARED_FOLDER_NAME = "Shared Folder Name";
    public static final String SHARED_FILE_NAME = "Shared File Name";
    public static final String KEY_ROWS = "Key Rows";
    public static final String LOOKUP_COLUMN = "Lookup Column";
    public static final String GET_ALL_RECORD = "Get All Record";
    public static final String COLUMNS = "columns";
    public static final String ROW = "Row";
    public static final String LOOKUP_ROWS = "Lookup Rows";
    public static final String LOOKUP_COLUMNS = "Lookup Columns";
    public static final String COLUMN_NUMBER = "Column Number";
    public static final String CELL_ID_DATA = "Cell ID Data";
    public static final String PATH = "Path";
    public static final String NAME = "Name";
    public static final String EXCEL_FILE = "Excel File";
    public static final String FILE_PATH = "File Path";
    public static final String COLUMN_RANGE = "Column Range";
    public static final String KEY_COLUMN_NAME = "Key Column Name";
    public static final String SEPARATOR = "Separator";
    public static final String AUTHORIZATION = "AUTHORIZATION";
    public static final String LOGIC_ERROR = "LOGIC_ERROR";

    public static String safeString(String value) {
        return value != null ? value : "";
    }

}
