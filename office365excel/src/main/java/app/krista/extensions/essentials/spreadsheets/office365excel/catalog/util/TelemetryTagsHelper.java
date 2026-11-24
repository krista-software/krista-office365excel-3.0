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

import app.krista.model.base.File;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryResources.*;

/**
 * Helper class for creating telemetry tags
 */
public class TelemetryTagsHelper {

    private TelemetryTagsHelper() {
    }

    private static boolean isNotEmpty(String str) {
        return str != null && !str.trim().isEmpty();
    }

    private static boolean isNotEmpty(List<?> list) {
        return list != null && !list.isEmpty();
    }

    private static boolean isRowsValid(List<String> rows) {
        return rows != null && !rows.isEmpty() && !rows.stream().allMatch(Objects::isNull);
    }

    private static String toCountString(List<?> list) {
        return list != null ? String.valueOf(list.size()) : "0";
    }

    private static String toLengthString(String str) {
        return str != null ? String.valueOf(str.length()) : "0";
    }

    // Builder class for common tag operations
    public static class TelemetryTagsBuilder {
        private final Map<String, String> tags = new HashMap<>();

        public TelemetryTagsBuilder operationType(String operationType) {
            tags.put(OPERATION_TYPE, operationType);
            return this;
        }

        public TelemetryTagsBuilder stringTag(String key, String value) {
            tags.put(key, String.valueOf(isNotEmpty(value)));
            return this;
        }

        public TelemetryTagsBuilder listTag(String hasKey, String countKey, List<?> list) {
            tags.put(hasKey, String.valueOf(isNotEmpty(list)));
            tags.put(countKey, toCountString(list));
            return this;
        }

        public TelemetryTagsBuilder rowsTag(List<String> rows) {
            tags.put(HAS_ROWS, String.valueOf(isRowsValid(rows)));
            tags.put(ROW_COUNT, toCountString(rows));
            return this;
        }

        public TelemetryTagsBuilder stringWithLength(String hasKey, String lengthKey, String value) {
            tags.put(hasKey, String.valueOf(isNotEmpty(value)));
            tags.put(lengthKey, toLengthString(value));
            return this;
        }

        public TelemetryTagsBuilder booleanTag(String key, Boolean value) {
            tags.put(key, String.valueOf(value));
            return this;
        }

        public TelemetryTagsBuilder nullableTag(String key, Object value) {
            tags.put(key, String.valueOf(value != null));
            return this;
        }

        public Map<String, String> build() {
            return new HashMap<>(tags);
        }
    }

    // Simplified methods using builder
    public static Map<String, String> createTableOperationTags(List<String> rows, List<File> rowsFromCsvFile) {
        return new TelemetryTagsBuilder()
                .operationType(TABLE_INSERT)
                .rowsTag(rows)
                .listTag(HAS_CSV_FILE, CSV_FILE_COUNT, rowsFromCsvFile)
                .build();
    }

    public static Map<String, String> createSheetOperationTags(List<String> rows, List<File> rowsFromFile) {
        return new TelemetryTagsBuilder()
                .operationType(SHEET_INSERT)
                .rowsTag(rows)
                .listTag(HAS_FILE, FILE_COUNT, rowsFromFile)
                .build();
    }

    public static Map<String, String> createRangeOperationTags(String range, String sheetName) {
        return new TelemetryTagsBuilder()
                .operationType(RANGE_SUM)
                .stringTag(HAS_RANGE, range)
                .stringTag(SHEET_PROVIDED, sheetName)
                .build();
    }

    public static Map<String, String> createVlookupOperationTags(String lookupValue, String range, String sheetName) {
        return new TelemetryTagsBuilder()
                .operationType(VLOOKUP)
                .stringTag(HAS_LOOKUP_VALUE, lookupValue)
                .stringTag(HAS_RANGE, range)
                .stringTag(SHEET_PROVIDED, sheetName)
                .build();
    }

    public static Map<String, String> createTableQueryOperationTags(String tableName) {
        return new TelemetryTagsBuilder()
                .operationType(TABLE_QUERY)
                .stringTag(HAS_TABLE_NAME, tableName)
                .build();
    }

    public static Map<String, String> createCellOperationTags(String cellId, String data) {
        return new TelemetryTagsBuilder()
                .operationType(CELL_UPDATE)
                .stringTag(HAS_CELL_ID, cellId)
                .stringWithLength(HAS_DATA, DATA_LENGTH, data)
                .build();
    }

    public static Map<String, String> createDeleteColumnOperationTags(String tableName, String columnName) {
        return new TelemetryTagsBuilder()
                .operationType(DELETE_COLUMN)
                .stringTag(HAS_TABLE_NAME, tableName)
                .stringTag(HAS_COLUMN_NAME, columnName)
                .build();
    }

    public static Map<String, String> createColumnValuesForKeysOperationTags(String tableName, List<String> keyRows, List<String> lookupColumn) {
        return new TelemetryTagsBuilder()
                .operationType(COLUMN_VALUES_FOR_KEYS)
                .stringTag(HAS_TABLE_NAME, tableName)
                .listTag(HAS_LOOKUP_ROWS, KEY_ROWS_COUNT, keyRows)
                .listTag(HAS_LOOKUP_COLUMNS, LOOKUP_COLUMN_COUNT, lookupColumn)
                .build();
    }

    public static Map<String, String> createExecuteFormulaOperationTags(String formula, String formulaName) {
        return new TelemetryTagsBuilder()
                .operationType(EXECUTE_FORMULA)
                .stringWithLength(HAS_FORMULA, FORMULA_LENGTH, formula)
                .stringTag(HAS_FORMULA_NAME, formulaName)
                .build();
    }

    public static Map<String, String> createVlookupInTableOperationTags(String lookupValue, String tableName, Double columnIndex) {
        return new TelemetryTagsBuilder()
                .operationType(VLOOKUP_IN_TABLE)
                .stringTag(HAS_LOOKUP_VALUE, lookupValue)
                .stringTag(HAS_TABLE_NAME, tableName)
                .nullableTag(HAS_COLUMN_INDEX, columnIndex)
                .build();
    }

    public static Map<String, String> createColumnsSumOperationTags(String tableName, List<String> columns) {
        return new TelemetryTagsBuilder()
                .operationType(COLUMNS_SUM)
                .stringTag(HAS_TABLE_NAME, tableName)
                .listTag(HAS_COLUMNS, COLUMN_COUNT, columns)
                .build();
    }

    public static Map<String, String> createSearchOperationTags(String sheetName, String dataToSearch) {
        return new TelemetryTagsBuilder()
                .operationType(SEARCH)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_DATA_TO_SEARCH, dataToSearch)
                .build();
    }

    public static Map<String, String> createCellDataOperationTags(String sheetName, String cellId) {
        return new TelemetryTagsBuilder()
                .operationType(CELL_DATA)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_CELL_ID, cellId)
                .build();
    }

    public static Map<String, String> createGetSharedExcelFileOperationTags(String sharedFolderName, String sharedFileName) {
        return new TelemetryTagsBuilder()
                .operationType(GET_SHARED_EXCEL_FILE)
                .stringTag(HAS_SHARED_FOLDER_NAME, sharedFolderName)
                .stringTag(HAS_SHARED_FILE_NAME, sharedFileName)
                .build();
    }

    public static Map<String, String> createConvertToPdfOperationTags(String driveId, String fileId) {
        return new TelemetryTagsBuilder()
                .operationType(CONVERT_TO_PDF)
                .stringTag(HAS_DRIVE_ID, driveId)
                .stringTag(HAS_FILE_ID, fileId)
                .build();
    }

    public static Map<String, String> createDeleteColumnWithIndexOperationTags(String sheetName, String columnName) {
        return new TelemetryTagsBuilder()
                .operationType(DELETE_COLUMN_WITH_INDEX)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_COLUMN_NAME, columnName)
                .build();
    }

    public static Map<String, String> createAddRowWithIndexOperationTags(String sheetName, List<String> rows, List<File> rowsFromCsv, Double rowIndex) {
        return new TelemetryTagsBuilder()
                .operationType(ADD_ROW_WITH_INDEX)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .rowsTag(rows)
                .listTag(HAS_CSV_FILE, CSV_FILE_COUNT, rowsFromCsv)
                .nullableTag(HAS_ROW_INDEX, rowIndex)
                .build();
    }

    public static Map<String, String> createUpdateRowWithIndexOperationTags(String sheetName, String row, Double rowIndex) {
        return new TelemetryTagsBuilder()
                .operationType(UPDATE_ROW_WITH_INDEX)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_ROW_DATA, row)
                .nullableTag(HAS_ROW_INDEX, rowIndex)
                .build();
    }

    public static Map<String, String> createGetCompleteRowInfoOperationTags(String sheetName, Double rowIndex) {
        return new TelemetryTagsBuilder()
                .operationType(GET_COMPLETE_ROW_INFO)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .nullableTag(HAS_ROW_INDEX, rowIndex)
                .build();
    }

    public static Map<String, String> createDeleteRowWithIndexOperationTags(String sheetName, Double rowIndex) {
        return new TelemetryTagsBuilder()
                .operationType(DELETE_ROW_WITH_INDEX)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .nullableTag(HAS_ROW_INDEX, rowIndex)
                .build();
    }

    public static Map<String, String> createAddMultipleCellDataOperationTags(String sheetName, List<Map<String, Object>> cellIDData) {
        return new TelemetryTagsBuilder()
                .operationType(ADD_MULTIPLE_CELL_DATA)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .listTag(HAS_CELL_DATA, CELL_DATA_COUNT, cellIDData)
                .build();
    }

    public static Map<String, String> createGetColumnValuesForKeysInSheetOperationTags(String sheetName, String range,
                                                                                       List<String> lookupRows, List<String> lookupColumns,
                                                                                       Boolean getAllRecords, Double columnNumber) {
        return new TelemetryTagsBuilder()
                .operationType(GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_RANGE, range)
                .listTag(HAS_LOOKUP_ROWS, LOOKUP_ROWS_COUNT, lookupRows)
                .listTag(HAS_LOOKUP_COLUMNS, LOOKUP_COLUMNS_COUNT, lookupColumns)
                .booleanTag(GET_ALL_RECORDS, getAllRecords)
                .nullableTag(HAS_COLUMN_NUMBER, columnNumber)
                .build();
    }

    public static Map<String, String> createClearSharedSheetForRangeOperationTags(String sheetName, String cellRange) {
        return new TelemetryTagsBuilder()
                .operationType(CLEAR_SHARED_SHEET_FOR_RANGE)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_CELL_RANGE, cellRange)
                .build();
    }

    public static Map<String, String> createDeleteSheetOperationTags(String sheetName) {
        return new TelemetryTagsBuilder()
                .operationType(DELETE_SHEET)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .build();
    }

    public static Map<String, String> createGetCellValueOperationTags(String fileId, String sheetName, String cellId) {
        return new TelemetryTagsBuilder()
                .operationType(GET_CELL_VALUE)
                .stringTag(HAS_FILE_ID, fileId)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_CELL_ID, cellId)
                .build();
    }

    public static Map<String, String> createAddRowsOperationTags(List<String> rows, List<File> rowsFromCsv) {
        return new TelemetryTagsBuilder()
                .operationType(ADD_ROWS)
                .rowsTag(rows)
                .listTag(HAS_CSV_FILE, CSV_FILE_COUNT, rowsFromCsv)
                .build();
    }

    public static Map<String, String> createFindValueOperationTags(String sheetName, String lookupValue) {
        return new TelemetryTagsBuilder()
                .operationType(FIND_VALUE)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_LOOKUP_VALUE, lookupValue)
                .build();
    }

    public static Map<String, String> createGetExcelFileInfoOperationTags(String path, String name) {
        return new TelemetryTagsBuilder()
                .operationType(GET_EXCEL_FILE_INFO)
                .stringTag(HAS_PATH, path)
                .stringTag(HAS_NAME, name)
                .build();
    }

    public static Map<String, String> createGetSheetsInfoOperationTags(String fileId) {
        return new TelemetryTagsBuilder()
                .operationType(GET_SHEETS_INFO)
                .stringTag(HAS_FILE_ID, fileId)
                .build();
    }

    public static Map<String, String> createAddDataToCellOperationTags(String cellID, String data) {
        return new TelemetryTagsBuilder()
                .operationType(ADD_DATA_TO_CELL)
                .stringTag(HAS_CELL_ID_UPPER, cellID)
                .stringTag(HAS_DATA, data)
                .build();
    }

    public static Map<String, String> createCreateSheetOperationTags(String newSheetName) {
        return new TelemetryTagsBuilder()
                .operationType(CREATE_SHEET)
                .stringTag(HAS_SHEET_NAME, newSheetName)
                .build();
    }

    public static Map<String, String> createClearSheetForRangeOperationTagsForSheet(String sheetName, String cellRange) {
        return new TelemetryTagsBuilder()
                .operationType(CLEAR_SHEET_FOR_RANGE)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_CELL_RANGE, cellRange)
                .build();
    }

    public static Map<String, String> createAddFormulaToCellOperationTags(String cellID, String formula) {
        return new TelemetryTagsBuilder()
                .operationType(ADD_FORMULA_TO_CELL)
                .stringTag(HAS_CELL_ID_UPPER, cellID)
                .stringTag(HAS_FORMULA, formula)
                .build();
    }

    public static Map<String, String> createMinMaxValueOperationTags(String keyColumnName, String columnRange) {
        Map<String, String> telemetryTags = new HashMap<>();
        telemetryTags.put(OPERATION_TYPE, GET_MIN_MAX_VALUE);
        telemetryTags.put(HAS_KEY_COLUMN_NAME, String.valueOf(keyColumnName != null && !keyColumnName.trim().isEmpty()));
        telemetryTags.put(HAS_COLUMN_RANGE, String.valueOf(columnRange != null && !columnRange.trim().isEmpty()));
        return telemetryTags;
    }

    public static Map<String, String> createUploadFileAndConvertToPDFOperationTags(File excelFile, String filePath) {
        Map<String, String> telemetryTags = new HashMap<>();
        telemetryTags.put(OPERATION_TYPE, UPLOAD_FILE_AND_CONVERT_TO_PDF);
        telemetryTags.put(HAS_EXCEL_FILE, String.valueOf(excelFile != null));
        telemetryTags.put(HAS_FILE_PATH, String.valueOf(filePath != null && !filePath.trim().isEmpty()));
        return telemetryTags;
    }

    public static Map<String, String> createTableVlookupOperationTags(String lookupValue, String tableName, Double columnIndex) {
        return new TelemetryTagsBuilder()
                .operationType(TABLE_VLOOKUP)
                .stringTag(HAS_LOOKUP_VALUE, lookupValue)
                .stringTag(HAS_TABLE_NAME, tableName)
                .nullableTag(HAS_COLUMN_INDEX, columnIndex)
                .build();
    }

    public static Map<String, String> createTableCreationOperationTags(String sheetName, String range) {
        return new TelemetryTagsBuilder()
                .operationType(TABLE_CREATION)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .stringTag(HAS_RANGE, range)
                .build();
    }

    public static Map<String, String> createTableWithIndexOperationTags(String tableName, List<String> rows, List<File> rowsFromCsv, Double index) {
        return new TelemetryTagsBuilder()
                .operationType(ADD_ROWS_TO_TABLE_WITH_INDEX)
                .stringTag(HAS_TABLE_NAME, tableName)
                .rowsTag(rows)
                .listTag(HAS_CSV_FILE, CSV_FILE_COUNT, rowsFromCsv)
                .nullableTag(HAS_INDEX, index)
                .build();
    }

    public static Map<String, String> createTableSummationOperationTags(String tableName, List<String> columns) {
        return new TelemetryTagsBuilder()
                .operationType(TABLE_SUMMATION)
                .stringTag(HAS_TABLE_NAME, tableName)
                .listTag(HAS_COLUMNS, COLUMN_COUNT, columns)
                .build();
    }

    public static Map<String, String> createAddRowsInTableOperationTags(String tableName, List<String> rows, List<File> rowsFromCsv) {
        return new TelemetryTagsBuilder()
                .operationType(ADD_ROWS_IN_TABLE)
                .stringTag(HAS_TABLE_NAME, tableName)
                .rowsTag(rows)
                .listTag(HAS_CSV_FILE, CSV_FILE_COUNT, rowsFromCsv)
                .build();
    }

    public static Map<String, String> createDeleteTableColumnOperationTags(String tableName, String columnName) {
        return new TelemetryTagsBuilder()
                .operationType(DELETE_TABLE_COLUMN)
                .stringTag(HAS_TABLE_NAME, tableName)
                .stringTag(HAS_COLUMN_NAME, columnName)
                .build();
    }

    public static Map<String, String> createGetColumnValuesForKeysInSheetOperationTags(String sheetName, List<String> lookupRows,
                                                                                       List<String> lookupColumns) {
        return new TelemetryTagsBuilder()
                .operationType(GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET)
                .stringTag(HAS_SHEET_NAME, sheetName)
                .listTag(HAS_LOOKUP_ROWS, LOOKUP_ROWS_COUNT, lookupRows)
                .listTag(HAS_LOOKUP_COLUMNS, LOOKUP_COLUMNS_COUNT, lookupColumns)
                .build();
    }

    public static Map<String, String> createGetColumnValuesForKeysOperationTags(String tableName, List<String> keyRows,
                                                                                List<String> lookupColumns) {
        return new TelemetryTagsBuilder()
                .operationType(GET_COLUMN_VALUES_FOR_KEYS_OPERATION)
                .stringTag(HAS_TABLE_NAME, tableName)
                .listTag(HAS_LOOKUP_ROWS, KEY_ROWS_COUNT, keyRows)
                .listTag(HAS_LOOKUP_COLUMNS, LOOKUP_COLUMN_COUNT, lookupColumns)
                .build();
    }

    public static Map<String, String> createUploadFileAndConvertToPDFOperationTags(String filePath) {
        return new TelemetryTagsBuilder()
                .operationType(UPLOAD_FILE_AND_CONVERT_TO_PDF_OPERATION)
                .stringTag(HAS_FILE_PATH, filePath)
                .build();
    }

}
