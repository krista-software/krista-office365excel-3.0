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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.builder;

import app.krista.model.base.File;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

public class StateMetaBuilder {

    private final Map<String, Object> metaData = new LinkedHashMap<>();

    private StateMetaBuilder() {
    }

    public static StateMetaBuilder builder() {
        return new StateMetaBuilder();
    }

    public StateMetaBuilder withDriveId(String driveId) {
        metaData.put(DRIVE_ID, driveId);
        return this;
    }

    public StateMetaBuilder withFileId(String fileId) {
        metaData.put(FILE_ID, fileId);
        return this;
    }

    public StateMetaBuilder withTableName(String tableName) {
        metaData.put(TABLE_NAME, tableName);
        return this;
    }

    public StateMetaBuilder withSheetName(String sheetName) {
        metaData.put(SHEET_NAME, sheetName);
        return this;
    }

    public StateMetaBuilder withRows(List<String> rows) {
        metaData.put(ROWS, rows);
        return this;
    }

    public StateMetaBuilder withRowsFromCsvFile(List<File> rowsFromCsvFile) {
        metaData.put(ROWS_FROM_CSV_FILE, rowsFromCsvFile);
        return this;
    }

    public StateMetaBuilder withStateId(String stateId) {
        metaData.put(STATE_ID, stateId);
        return this;
    }

    public StateMetaBuilder withRange(String range) {
        metaData.put(RANGE, range);
        return this;
    }

    public StateMetaBuilder withLookupValue(String lookupValue) {
        metaData.put(LOOKUP_VALUE, lookupValue);
        return this;
    }

    public StateMetaBuilder withColumnIndex(Double columnIndex) {
        metaData.put(COLUMN_INDEX, columnIndex);
        return this;
    }

    public StateMetaBuilder withExactMatch(Boolean exactMatch) {
        metaData.put(EXACT_MATCH, exactMatch);
        return this;
    }

    public StateMetaBuilder withCellId(String cellId) {
        metaData.put(CELL_ID, cellId);
        return this;
    }

    public StateMetaBuilder withData(String data) {
        metaData.put(DATA, data);
        return this;
    }

    public StateMetaBuilder withRowIndex(Double rowIndex) {
        metaData.put(ROW_INDEX, rowIndex);
        return this;
    }

    public StateMetaBuilder withColumnName(String columnName) {
        metaData.put(COLUMN_NAME, columnName);
        return this;
    }

    public StateMetaBuilder withFormula(String formula) {
        metaData.put(FORMULA, formula);
        return this;
    }

    public StateMetaBuilder withFormulaName(String formulaName) {
        metaData.put(FORMULA_NAME, formulaName);
        return this;
    }

    public StateMetaBuilder withComment(String comment) {
        metaData.put(COMMENT, comment);
        return this;
    }

    public StateMetaBuilder withDataToSearch(String dataToSearch) {
        metaData.put(DATA_TO_SEARCH, dataToSearch);
        return this;
    }

    public StateMetaBuilder withKeyRows(List<String> keyRows) {
        metaData.put(KEY_ROWS, keyRows);
        return this;
    }

    public StateMetaBuilder withLookupColumn(List<String> lookupColumn) {
        metaData.put(LOOKUP_COLUMN, lookupColumn);
        return this;
    }

    public StateMetaBuilder withGetAllRecord(Boolean getAllRecord) {
        metaData.put(GET_ALL_RECORD, getAllRecord);
        return this;
    }

    public StateMetaBuilder withSharedFolderName(String sharedFolderName) {
        metaData.put(SHARED_FOLDER_NAME, sharedFolderName);
        return this;
    }

    public StateMetaBuilder withSharedFileName(String sharedFileName) {
        metaData.put(SHARED_FILE_NAME, sharedFileName);
        return this;
    }

    public StateMetaBuilder withPath(String path) {
        metaData.put(PATH, path);
        return this;
    }

    public StateMetaBuilder withName(String name) {
        metaData.put(NAME, name);
        return this;
    }

    public StateMetaBuilder withExcelFile(File excelFile) {
        metaData.put(EXCEL_FILE, excelFile);
        return this;
    }

    public StateMetaBuilder withFilePath(String filePath) {
        metaData.put(FILE_PATH, filePath);
        return this;
    }

    public StateMetaBuilder withKeyColumnName(String keyColumnName) {
        metaData.put(KEY_COLUMN_NAME, keyColumnName);
        return this;
    }

    public StateMetaBuilder withColumnRange(String columnRange) {
        metaData.put(COLUMN_RANGE, columnRange);
        return this;
    }

    public StateMetaBuilder withSeparator(String separator) {
        metaData.put(SEPARATOR, separator);
        return this;
    }

    public StateMetaBuilder withColumns(List<String> columns) {
        metaData.put(COLUMNS, columns);
        return this;
    }

    public StateMetaBuilder withColumnNumber(Double columnNumber) {
        metaData.put(COLUMN_NUMBER, columnNumber);
        return this;
    }

    public StateMetaBuilder withAdditional(String key, Object value) {
        metaData.put(key, value);
        return this;
    }

    public StateMetaBuilder withLookupRows(List<String> lookupRows) {
        metaData.put(LOOKUP_ROWS, lookupRows);
        return this;
    }

    public StateMetaBuilder withLookupColumns(List<String> lookupColumns) {
        metaData.put(LOOKUP_COLUMNS, lookupColumns);
        return this;
    }

    public Map<String, Object> build() {
        return metaData;
    }
}
