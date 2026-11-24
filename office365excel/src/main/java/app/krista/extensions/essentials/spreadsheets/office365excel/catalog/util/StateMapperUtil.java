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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.builder.StateMetaBuilder;
import app.krista.model.base.File;
import org.jetbrains.annotations.NotNull;

import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

/**
 * Utility class for creating metadata maps used in state management
 * between catalog and subcatalog requests for Office365Excel extension.
 */
public class StateMapperUtil {

    private StateMapperUtil() {
    }

    /**
     * Creates a metadata map for adding rows to a table operation
     */
    @NotNull
    public static Map<String, Object> addRowsToTableMetaToMap(String driveId, String fileId, String tableName,
                                                              List<String> rows, List<File> rowsFromCsvFile, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withTableName(tableName)
                .withRows(rows)
                .withRowsFromCsvFile(rowsFromCsvFile)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for adding rows to a sheet operation
     */
    @NotNull
    public static Map<String, Object> addRowsToSheetMetaToMap(String driveId, String fileId, String sheetName,
                                                              List<String> rows, List<File> rowsFromCsvFile, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRows(rows)
                .withRowsFromCsvFile(rowsFromCsvFile)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for range sum operation
     */
    @NotNull
    public static Map<String, Object> rangeSumMetaToMap(String driveId, String fileId, String sheetName,
                                                        String range, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRange(range)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get table rows operation
     */
    @NotNull
    public static Map<String, Object> getTableRowsMetaToMap(String driveId, String fileId, String tableName,
                                                            String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withTableName(tableName)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for sheet vlookup operation
     */
    @NotNull
    public static Map<String, Object> sheetVlookupMetaToMap(String driveId, String fileId, String lookupValue,
                                                            String sheetName, String range, Double columnIndex,
                                                            Boolean exactMatch, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withLookupValue(lookupValue)
                .withSheetName(sheetName)
                .withRange(range)
                .withColumnIndex(columnIndex)
                .withExactMatch(exactMatch)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a map with metadata for add data to cell operation
     */
    @NotNull
    public static Map<String, Object> addDataToCellMetaToMap(String driveId, String fileId, String sheetName,
                                                             String cellId, String data, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withCellId(cellId)
                .withData(data)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get rows for range operation
     */
    @NotNull
    public static Map<String, Object> getRowsForRangeMetaToMap(String driveId, String fileId, String sheetName,
                                                               String range, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRange(range)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for vlookup in table operation
     */
    @NotNull
    public static Map<String, Object> vlookupInTableMetaToMap(String driveId, String fileId, String lookupValue,
                                                              String tableName, Double columnIndex, Boolean exactMatch, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withLookupValue(lookupValue)
                .withTableName(tableName)
                .withColumnIndex(columnIndex)
                .withExactMatch(exactMatch)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for delete row with index operation
     */
    @NotNull
    public static Map<String, Object> deleteRowWithIndexMetaToMap(String driveId, String fileId, String sheetName,
                                                                  Double rowIndex, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRowIndex(rowIndex)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get complete row info operation
     */
    @NotNull
    public static Map<String, Object> getCompleteRowInfoMetaToMap(String driveId, String fileId, String sheetName,
                                                                  Double rowIndex, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRowIndex(rowIndex)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for delete column operation
     */
    @NotNull
    public static Map<String, Object> deleteColumnMetaToMap(String driveId, String fileId, String tableName,
                                                            String columnName, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withTableName(tableName)
                .withColumnName(columnName)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for executeFormula operation
     */
    public static Map<String, Object> executeFormulaMetaToMap(String driveId, String fileId, String formula,
                                                              String formulaName, String comment, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withFormula(formula)
                .withFormulaName(formulaName)
                .withComment(comment)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for search operation
     */
    @NotNull
    public static Map<String, Object> searchMetaToMap(String driveId, String fileId, String sheetName,
                                                      String dataToSearch, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withDataToSearch(dataToSearch)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for cell data operation
     */
    @NotNull
    public static Map<String, Object> cellDataMetaToMap(String driveId, String fileId, String sheetName,
                                                        String cellId, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withCellId(cellId)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get shared Excel file operation
     */
    @NotNull
    public static Map<String, Object> getSharedExcelFileMetaToMap(String sharedFolderName, String sharedFileName, String stateId) {
        return StateMetaBuilder.builder()
                .withSharedFolderName(sharedFolderName)
                .withSharedFileName(sharedFileName)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get column values for keys operation
     */
    @NotNull
    public static Map<String, Object> getColumnValuesForKeysMetaToMap(String driveId, String fileId, String tableName,
                                                                      List<String> keyRows, List<String> lookupColumn,
                                                                      Boolean getAllRecord, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withTableName(tableName)
                .withKeyRows(keyRows)
                .withLookupColumn(lookupColumn)
                .withGetAllRecord(getAllRecord)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for columns sum operation
     */
    @NotNull
    public static Map<String, Object> columnsSumMetaToMap(String driveId, String fileId, String tableName,
                                                          List<String> columns, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withTableName(tableName)
                .withColumns(columns)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for convert to PDF operation
     */
    @NotNull
    public static Map<String, Object> convertToPdfMetaToMap(String driveId, String fileId, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for delete column with index operation
     */
    @NotNull
    public static Map<String, Object> deleteColumnWithIndexMetaToMap(String driveId, String fileId, String sheetName,
                                                                     String columnName, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withColumnName(columnName)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for add row with index operation
     */
    @NotNull
    public static Map<String, Object> addRowWithIndexMetaToMap(String driveId, String fileId, String sheetName,
                                                               List<String> rows, List<File> rowsFromCsv,
                                                               Double rowIndex, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRows(rows)
                .withRowsFromCsvFile(rowsFromCsv)
                .withRowIndex(rowIndex)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for update row with index operation
     */
    @NotNull
    public static Map<String, Object> updateRowWithIndexMetaToMap(String driveId, String fileId, String sheetName,
                                                                  String row, Double rowIndex, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withAdditional(ROW, row)
                .withRowIndex(rowIndex)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get column values for keys in sheet operation
     */
    @NotNull
    public static Map<String, Object> getColumnValuesForKeysInSheetMetaToMap(String driveId, String fileId, String sheetName,
                                                                             String range, List<String> lookupRows,
                                                                             List<String> lookupColumns, Boolean getAllRecords,
                                                                             Double columnNumber, String stateId) {
        StateMetaBuilder builder = StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRange(range)
                .withLookupRows(lookupRows)
                .withLookupColumns(lookupColumns)
                .withGetAllRecord(getAllRecords)
                .withStateId(stateId);
        if (columnNumber != null) builder.withColumnNumber(columnNumber);
        return builder.build();
    }

    /**
     * Creates a map with metadata for add multiple cell data operation
     */
    @NotNull
    public static Map<String, Object> addMultipleCellDataMetaToMap(String driveId, String fileId, String sheetName,
                                                                   List<Map<String, Object>> cellIDData, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withAdditional(CELL_ID_DATA, cellIDData)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a map with metadata for add multiple cell data operation for sheet
     */
    @NotNull
    public static Map<String, Object> addMultipleCellDataMetaToMapForSheet(String fileId, String sheetName,
                                                                   List<Map<String, Object>> cellIDData, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withAdditional(CELL_ID_DATA, cellIDData)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a map with metadata for clearing a shared sheet range
     */
    @NotNull
    public static Map<String, Object> clearSharedSheetMetaToMap(String driveId, String fileId, String sheetName,
                                                                String cellRange, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRange(cellRange)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for delete sheet operation
     */
    @NotNull
    public static Map<String, Object> deleteSheetMetaToMap(String fileId, String sheetName, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withStateId(stateId)
                .build();
    }

    @NotNull
    public static Map<String, Object> getCellValueMetaToMap(String fileId, String sheetName, String cellId, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withCellId(cellId)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for adding rows to a sheet operation
     */
    @NotNull
    public static Map<String, Object> addRowsMetaToMap(String fileId, String sheetName,
                                                       List<String> rows, List<File> rowsFromCsvFile, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRows(rows)
                .withRowsFromCsvFile(rowsFromCsvFile)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for adding a row at a specific index operation
     */
    @NotNull
    public static Map<String, Object> addRowWithIndexMetaToMap(String fileId, String sheetName,
                                                               List<String> rows, List<File> rowsFromCsvFile,
                                                               Double rowIndex, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRows(rows)
                .withRowsFromCsvFile(rowsFromCsvFile)
                .withRowIndex(rowIndex)
                .withStateId(stateId)
                .build();
    }

    @NotNull
    public static Map<String, Object> findValueMetaToMap(String fileId, String sheetName, String lookupValue, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withLookupValue(lookupValue)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a map with metadata for Excel file information operation
     */
    @NotNull
    public static Map<String, Object> getExcelFileInfoMetaToMap(String path, String name, String stateId) {
        return StateMetaBuilder.builder()
                .withPath(path)
                .withName(name)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a map with metadata for sheets information operation
     */
    @NotNull
    public static Map<String, Object> getSheetsInfoMetaToMap(String fileId, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a map with metadata for create sheet operation
     */
    @NotNull
    public static Map<String, Object> createSheetMetaToMap(String fileId, String sheetName, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a map with metadata for clear sheet for range operation
     */
    @NotNull
    public static Map<String, Object> clearSheetForRangeMetaToMap(String fileId, String sheetName, String range, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRange(range)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for vlookup operation
     */
    @NotNull
    public static Map<String, Object> vlookupMetaToMap(String fileId, String lookupValue, String sheetName,
                                                       String range, Double columnIndex, Boolean exactMatch, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withLookupValue(lookupValue)
                .withSheetName(sheetName)
                .withRange(range)
                .withColumnIndex(columnIndex)
                .withExactMatch(exactMatch != null ? exactMatch : false)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for add formula to cell operation
     */
    @NotNull
    public static Map<String, Object> addFormulaToCellMetaToMap(String fileId, String sheetName,
                                                                String cellId, String formula, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withCellId(cellId)
                .withFormula(formula)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for summation operation
     */
    @NotNull
    public static Map<String, Object> summationMetaToMap(String fileId, String sheetName, String range, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRange(range)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for upload file and convert to PDF operation
     */
    @NotNull
    public static Map<String, Object> uploadFileAndConvertToPDFMetaToMap(File excelFile, String filePath, String stateId) {
        return StateMetaBuilder.builder()
                .withExcelFile(excelFile)
                .withFilePath(filePath)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get min-max value operation
     */
    @NotNull
    public static Map<String, Object> getMinMaxValueMetaToMap(String fileId, String sheetName,
                                                              String keyColumnName, String columnRange, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withKeyColumnName(keyColumnName)
                .withColumnRange(columnRange)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for vlookup in table operation
     */
    @NotNull
    public static Map<String, Object> vlookupTableMetaToMap(String fileId, String lookupValue, String tableName,
                                                            Double columnIndex, Boolean exactMatch, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withLookupValue(lookupValue)
                .withTableName(tableName)
                .withColumnIndex(columnIndex)
                .withExactMatch(exactMatch != null ? exactMatch : false)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for add rows to table operation
     */
    @NotNull
    public static Map<String, Object> addTableRowsMetaToMap(String fileId, String tableName,
                                                            List<String> rows, List<File> rowsFromCsvFile, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withTableName(tableName)
                .withRows(rows)
                .withRowsFromCsvFile(rowsFromCsvFile)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for add rows to table with index operation
     */
    @NotNull
    public static Map<String, Object> addRowsToTableWithIndexMetaToMap(String driveId, String fileId, String tableName,
                                                                       List<String> rows, List<File> rowsFromCsv,
                                                                       Double index, String separator, String stateId) {
        return StateMetaBuilder.builder()
                .withDriveId(driveId)
                .withFileId(fileId)
                .withTableName(tableName)
                .withRows(rows)
                .withRowsFromCsvFile(rowsFromCsv)
                .withRowIndex(index)
                .withAdditional(SEPARATOR, separator)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get rows operation
     */
    @NotNull
    public static Map<String, Object> getRowsMetaToMap(String workbookId, String tableName, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(workbookId)
                .withTableName(tableName)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for table summation operation
     */
    @NotNull
    public static Map<String, Object> tableSummationMetaToMap(String fileId, String tableName, List<String> columns, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withTableName(tableName)
                .withColumns(columns)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for delete table column operation
     */
    @NotNull
    public static Map<String, Object> deleteTableColumnMetaToMap(String fileId, String tableName, String columnName, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withTableName(tableName)
                .withColumnName(columnName)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for create table for range operation
     */
    @NotNull
    public static Map<String, Object> createTableForRangeMetaToMap(String fileId, String sheetName, String range,
                                                                   String driveId, String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withSheetName(sheetName)
                .withRange(range)
                .withDriveId(driveId)
                .withStateId(stateId)
                .build();
    }

    /**
     * Creates a metadata map for get column values for keys operation
     */
    @NotNull
    public static Map<String, Object> getTableColumnValuesForKeysMetaToMap(String fileId, String tableName,
                                                                           List<String> keyRows, List<String> lookupColumns,
                                                                           Boolean getAllRecord, Double columnNumber,
                                                                           String stateId) {
        return StateMetaBuilder.builder()
                .withFileId(fileId)
                .withTableName(tableName)
                .withKeyRows(keyRows)
                .withLookupColumns(lookupColumns)
                .withGetAllRecord(getAllRecord)
                .withColumnNumber(columnNumber)
                .withStateId(stateId)
                .build();
    }
}
