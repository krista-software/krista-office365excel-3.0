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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.model.base.File;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;

/**
 * Helper class for creating validation maps using Builder pattern
 */
public class ValidationMapHelper {

    /**
     * Builder class for creating validation maps
     */
    public static class ValidationMapBuilder {
        private final Map<Validator.ValidationResource, String> validationMap = new HashMap<>();

        public ValidationMapBuilder driveId(String driveId) {
            validationMap.put(Validator.ValidationResource.DRIVE_ID, driveId != null ? driveId : "");
            return this;
        }

        public ValidationMapBuilder fileId(String fileId) {
            validationMap.put(Validator.ValidationResource.FILE_ID, fileId != null ? fileId : "");
            return this;
        }

        public ValidationMapBuilder sheetName(String sheetName) {
            validationMap.put(Validator.ValidationResource.SHEET_NAME, sheetName != null ? sheetName : "");
            return this;
        }

        public ValidationMapBuilder tableName(String tableName) {
            validationMap.put(Validator.ValidationResource.TABLE_NAME, tableName != null ? tableName : "");
            return this;
        }

        public ValidationMapBuilder range(String range) {
            validationMap.put(Validator.ValidationResource.RANGE, range != null ? range : "");
            return this;
        }

        public ValidationMapBuilder cellId(String cellId) {
            validationMap.put(Validator.ValidationResource.CELL_ID, cellId != null ? cellId : "");
            return this;
        }

        public ValidationMapBuilder columnIndex(Double columnIndex) {
            validationMap.put(Validator.ValidationResource.COLUMN_INDEX, columnIndex != null ? columnIndex.toString() : "");
            return this;
        }

        public ValidationMapBuilder lookupValue(String lookupValue) {
            validationMap.put(Validator.ValidationResource.LOOKUP_VALUE, lookupValue != null ? lookupValue : "");
            return this;
        }

        public ValidationMapBuilder rowIndex(Double rowIndex) {
            validationMap.put(Validator.ValidationResource.ROW_INDEX, rowIndex != null ? rowIndex.toString() : "");
            return this;
        }

        public ValidationMapBuilder formula(String formula) {
            validationMap.put(Validator.ValidationResource.FORMULA, formula != null ? formula : "");
            return this;
        }

        public ValidationMapBuilder formulaName(String formulaName) {
            validationMap.put(Validator.ValidationResource.FORMULA_NAME, formulaName != null ? formulaName : "");
            return this;
        }

        public ValidationMapBuilder comment(String comment) {
            validationMap.put(Validator.ValidationResource.COMMENT, comment != null ? comment : "");
            return this;
        }

        public ValidationMapBuilder sharedFileName(String sharedFileName) {
            validationMap.put(Validator.ValidationResource.SHARED_FILE_NAME, sharedFileName != null ? sharedFileName : "");
            return this;
        }

        public ValidationMapBuilder conditionalRowData(List<String> rows, List<File> csvFiles) {
            boolean isRowsMissing = rows == null || rows.isEmpty() || rows.stream().allMatch(Objects::isNull);
            boolean isCsvFileMissing = csvFiles == null || csvFiles.isEmpty();

            if (isRowsMissing && isCsvFileMissing) {
                validationMap.put(Validator.ValidationResource.ROW_DATA, "");
                validationMap.put(Validator.ValidationResource.CSV_FILE, "");
            }
            return this;
        }

        public ValidationMapBuilder conditionalKeyRows(List<String> rows, List<File> csvFiles) {
            boolean isRowsMissing = rows == null || rows.isEmpty() || rows.stream().allMatch(Objects::isNull);
            boolean isCsvFileMissing = csvFiles == null || csvFiles.isEmpty();

            if (isRowsMissing && isCsvFileMissing) {
                validationMap.put(Validator.ValidationResource.KEY_ROWS, "");
                validationMap.put(Validator.ValidationResource.CSV_FILE, "");
            }
            return this;
        }

        public ValidationMapBuilder conditionalColumns(List<String> columns) {
            boolean isColumnsMissing = columns == null || columns.isEmpty() || columns.stream().allMatch(Objects::isNull);
            if (isColumnsMissing) {
                validationMap.put(Validator.ValidationResource.COLUMNS, "");
            }
            return this;
        }

        public ValidationMapBuilder conditionalLookupRows(List<String> keyRows) {
            if (keyRows == null || keyRows.isEmpty()) {
                validationMap.put(Validator.ValidationResource.LOOKUP_ROWS, "");
            }
            return this;
        }

        public ValidationMapBuilder custom(Consumer<ValidationMapBuilder> customizer) {
            customizer.accept(this);
            return this;
        }

        public ValidationMapBuilder columnName(String columnName) {
            validationMap.put(Validator.ValidationResource.COLUMN_NAME, columnName != null ? columnName : "");
            return this;
        }

        public ValidationMapBuilder rowData(String rowData) {
            validationMap.put(Validator.ValidationResource.ROW_DATA, rowData != null ? rowData : "");
            return this;
        }

        public Map<Validator.ValidationResource, String> build() {
            return new HashMap<>(validationMap);
        }
    }

    public static ValidationMapBuilder builder() {
        return new ValidationMapBuilder();
    }

    // Refactored methods using the builder pattern
    public static Map<Validator.ValidationResource, String> createAddRowsToTableValidationMap(String driveId, String fileId,
                                                                                              String tableName, List<String> rows,
                                                                                              List<File> rowsFromCsvFile) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .tableName(tableName)
                .conditionalRowData(rows, rowsFromCsvFile)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddRowsToTableValidationMapForTable(String fileId, String tableName,
                                                                                                      List<String> rows,
                                                                                                      List<File> rowsFromCsvFile, Double index) {
        return builder()
                .fileId(fileId)
                .tableName(tableName)
                .rowIndex(index)
                .conditionalKeyRows(rows, rowsFromCsvFile)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddRowsToSheetValidationMap(String driveId, String fileId,
                                                                                              String sheetName, List<String> rows,
                                                                                              List<File> rowsFromFile) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .conditionalRowData(rows, rowsFromFile)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createRangeValidationMap(String driveId, String fileId,
                                                                                     String sheetName, String range) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .range(range)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createVlookupValidationMap(String driveId, String fileId, String sheetName,
                                                                                       String range, Double columnIndex) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .range(range)
                .columnIndex(columnIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createColumnsSumValidationMap(String driveId, String fileId,
                                                                                          String tableName, List<String> columns) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .tableName(tableName)
                .conditionalColumns(columns)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createVlookupInTableValidationMap(String driveId, String fileId,
                                                                                              String tableName, String lookupValue,
                                                                                              Double columnIndex) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .tableName(tableName)
                .lookupValue(lookupValue)
                .columnIndex(columnIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createTableQueryValidationMap(String driveId, String fileId,
                                                                                          String tableName) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .tableName(tableName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createTableQueryValidationMapForTable(String fileId, String tableName) {
        return builder()
                .fileId(fileId)
                .tableName(tableName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createCellValidationMap(String driveId, String fileId, String sheetName,
                                                                                    String cellId) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .cellId(cellId)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createDeleteColumnValidationMap(String driveId, String fileId, String tableName,
                                                                                            String columnName) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .tableName(tableName)
                .columnName(columnName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createDeleteColumnWithIndexValidationMap(String driveId, String fileId,
                                                                                                     String sheetName, String columnName) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .columnName(columnName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createColumnValuesForKeysValidationMap(String driveId, String fileId,
                                                                                                   String tableName, List<String> keyRows,
                                                                                                   List<String> lookupColumn) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .tableName(tableName)
                .conditionalLookupRows(keyRows)
                .custom(b -> {
                    boolean isLookupColumnMissing = lookupColumn == null || lookupColumn.isEmpty() || lookupColumn.stream().allMatch(Objects::isNull);
                    if (isLookupColumnMissing) {
                        b.validationMap.put(Validator.ValidationResource.LOOKUP_COLUMNS, "");
                    }
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createColumnValuesForKeysValidationMapForTable(String fileId, String tableName,
                                                                                                           List<String> keyRows,
                                                                                                           List<String> lookupColumn) {
        return builder()
                .fileId(fileId)
                .tableName(tableName)
                .conditionalLookupRows(keyRows)
                .custom(b -> {
                    boolean isLookupColumnMissing = lookupColumn == null || lookupColumn.isEmpty() || lookupColumn.stream().allMatch(Objects::isNull);
                    if (isLookupColumnMissing) {
                        b.validationMap.put(Validator.ValidationResource.LOOKUP_COLUMNS, "");
                    }
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createExecuteFormulaValidationMap(String driveId, String fileId,
                                                                                              String formula, String formulaName,
                                                                                              String comment) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .formula(formula)
                .formulaName(formulaName)
                .comment(comment)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createSearchValidationMap(String driveId, String fileId,
                                                                                      String sheetName) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createCellDataValidationMap(String driveId, String fileId,
                                                                                        String sheetName, String cellId) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .cellId(cellId)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetSharedExcelFileValidationMap(String sharedFileName) {
        return builder()
                .sharedFileName(sharedFileName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createConvertToPdfValidationMap(String driveId, String fileId) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createConvertToPdfValidationMapForSheet(String fileId) {
        return builder()
                .fileId(fileId)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddRowWithIndexValidationMap(String driveId, String fileId,
                                                                                               String sheetName, List<String> rows,
                                                                                               List<File> rowsFromCsv, Double rowIndex) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .rowIndex(rowIndex)
                .conditionalRowData(rows, rowsFromCsv)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createUpdateRowWithIndexValidationMap(String driveId, String fileId,
                                                                                                  String sheetName, String row,
                                                                                                  Double rowIndex) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .custom(b -> b.validationMap.put(Validator.ValidationResource.ROW_DATA, row != null ? row : ""))
                .rowIndex(rowIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetCompleteRowInfoValidationMap(String driveId, String fileId,
                                                                                                  String sheetName, Double rowIndex) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .rowIndex(rowIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetCompleteRowInfoValidationMapForSheet(String fileId, String sheetName,
                                                                                                          Double rowIndex) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .rowIndex(rowIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createDeleteRowWithIndexValidationMap(String driveId, String fileId,
                                                                                                  String sheetName, Double rowIndex) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .rowIndex(rowIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddMultipleCellDataValidationMap(String driveId, String fileId,
                                                                                                   String sheetName) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddMultipleCellDataValidationMapForSheet(String fileId, String sheetName) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetColumnValuesForKeysInSheetValidationMap(String driveId, String fileId,
                                                                                                             String sheetName, String range,
                                                                                                             List<String> lookupRows, List<String> lookupColumns) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .range(range)
                .custom(b -> {
                    if (lookupRows == null || lookupRows.isEmpty()) {
                        b.validationMap.put(Validator.ValidationResource.LOOKUP_ROWS, "");
                    }
                    if (lookupColumns == null || lookupColumns.isEmpty()) {
                        b.validationMap.put(Validator.ValidationResource.LOOKUP_COLUMNS, "");
                    }
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetColumnValuesForKeysInSheetValidationMapForSheet(String fileId, String sheetName,
                                                                                                                     String range, List<String> keyRows,
                                                                                                                     List<String> keyColumns) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .range(range)
                .custom(b -> {
                    if (keyRows == null || keyRows.isEmpty()) {
                        b.validationMap.put(Validator.ValidationResource.LOOKUP_ROWS, "");
                    }
                    if (keyColumns == null || keyColumns.isEmpty()) {
                        b.validationMap.put(Validator.ValidationResource.LOOKUP_COLUMNS, "");
                    }
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createClearSharedSheetForRangeValidationMap(String driveId, String fileId,
                                                                                                        String sheetName, String cellRange) {
        return builder()
                .driveId(driveId)
                .fileId(fileId)
                .sheetName(sheetName)
                .range(cellRange)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createDeleteSheetValidationMap(String fileId, String sheetName) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetCellValueValidationMap(String fileId, String sheetName, String cellId) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .cellId(cellId)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddRowsValidationMap(String fileId, String sheetName, List<String> rows,
                                                                                       List<File> rowsFromCsv) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .conditionalRowData(rows, rowsFromCsv)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddRowWithIndexValidationMapForSheet(String fileId, String sheetName,
                                                                                                       List<String> rows, List<File> rowsFromCsv,
                                                                                                       Double rowIndex) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .rowIndex(rowIndex)
                .conditionalRowData(rows, rowsFromCsv)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createUpdateRowWithIndexValidationMapForSheet(String fileId, String sheetName,
                                                                                                          String row, Double rowIndex) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .rowData(row)
                .rowIndex(rowIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createDeleteColumnWithIndexValidationMapForSheet(String fileId, String sheetName,
                                                                                                             String columnName) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .columnName(columnName)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createDeleteRowWithIndexValidationMapForSheet(String fileId, String sheetName,
                                                                                                          Double rowIndex) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .rowIndex(rowIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createFindValueValidationMapForSheet(String fileId, String sheetName,
                                                                                                 String lookupValue) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .lookupValue(lookupValue)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetExcelFileInfoValidationMap(String path, String name) {
        return builder()
                .custom(b -> {
                    b.validationMap.put(Validator.ValidationResource.PATH, path != null ? path : "");
                    b.validationMap.put(Validator.ValidationResource.NAME, name != null ? name : "");
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createGetSheetsInfoValidationMap(String fileId) {
        return builder()
                .fileId(fileId)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddDataToCellValidationMap(String fileID, String sheetName,
                                                                                             String cellID) {
        return builder()
                .fileId(fileID)
                .sheetName(sheetName)
                .cellId(cellID)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createExecuteFormulaValidationMapForSheet(String fileId, String formula,
                                                                                                      String formulaName, String comment) {
        return builder()
                .fileId(fileId)
                .formula(formula)
                .formulaName(formulaName)
                .comment(comment)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createCreateSheetValidationMap(String fileId) {
        return builder()
                .fileId(fileId)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createClearSheetForRangeValidationMap(String fileID, String sheetName,
                                                                                                  String cellRange) {
        return builder()
                .fileId(fileID)
                .sheetName(sheetName)
                .range(cellRange)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createVlookupValidationMapForSheet(String fileId, String sheetName,
                                                                                               String range, Double columnIndex,
                                                                                               String lookupValue) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .range(range)
                .columnIndex(columnIndex)
                .lookupValue(lookupValue)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createRangeValidationMapForSheet(String fileId, String sheetName, String range) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .range(range)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddFormulaToCellValidationMap(String fileID, String sheetName,
                                                                                                String cellID, String formula) {
        return builder()
                .fileId(fileID)
                .sheetName(sheetName)
                .cellId(cellID)
                .formula(formula)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createMinMaxValueValidationMap(String fileId, String sheetName,
                                                                                           String keyColumnName, String columnRange) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .custom(b -> {
                    b.validationMap.put(Validator.ValidationResource.KEY_COLUMN_NAME, keyColumnName != null ? keyColumnName : "");
                    b.validationMap.put(Validator.ValidationResource.COLUMN_RANGE, columnRange != null ? columnRange : "");
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createUploadFileAndConvertToPDFValidationMap(File excelFile) {
        return builder()
                .custom(b -> {
                    boolean isExcelFileMissing = excelFile == null;
                    if (isExcelFileMissing) {
                        b.validationMap.put(Validator.ValidationResource.EXCEL_FILE, "");
                    }
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createTableVlookupValidationMap(String fileId, String lookupValue,
                                                                                            String tableName, Double columnIndex) {
        return builder()
                .fileId(fileId)
                .lookupValue(lookupValue)
                .tableName(tableName)
                .columnIndex(columnIndex)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createTableCreationValidationMap(String fileId, String sheetName,
                                                                                             String range) {
        return builder()
                .fileId(fileId)
                .sheetName(sheetName)
                .range(range)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createTableSummationValidationMap(String fileId, String tableName,
                                                                                              List<String> columns) {
        return builder()
                .fileId(fileId)
                .tableName(tableName)
                .conditionalColumns(columns)
                .build();
    }

    public static Map<Validator.ValidationResource, String> createAddRowsInTableValidationMap(String fileId, String tableName,
                                                                                              List<String> rows, List<File> rowsFromCsv) {
        return builder()
                .fileId(fileId)
                .tableName(tableName)
                .custom(b -> {
                    boolean isRowsMissing = rows == null || rows.isEmpty() || rows.stream().allMatch(Objects::isNull);
                    boolean isCsvFileMissing = rowsFromCsv == null || rowsFromCsv.isEmpty();
                    if (isRowsMissing && isCsvFileMissing) {
                        b.validationMap.put(Validator.ValidationResource.KEY_ROWS, "");
                        b.validationMap.put(Validator.ValidationResource.CSV_FILE, "");
                    }
                })
                .build();
    }

    public static Map<Validator.ValidationResource, String> createDeleteTableColumnValidationMap(String fileId, String tableName,
                                                                                                 String columnName) {
        return builder()
                .fileId(fileId)
                .tableName(tableName)
                .columnName(columnName)
                .build();
    }
}
