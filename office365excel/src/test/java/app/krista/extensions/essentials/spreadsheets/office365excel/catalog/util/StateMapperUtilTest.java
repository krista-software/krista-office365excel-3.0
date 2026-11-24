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
import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.mock;

class StateMapperUtilTest {

    @Test
    void testAddRowsToTableMetaToMap() {
        File mockFile = mock(File.class);
        List<String> rows = Arrays.asList("row1", "row2");
        List<File> csvFiles = Collections.singletonList(mockFile);

        Map<String, Object> result = StateMapperUtil.addRowsToTableMetaToMap("drive1", "file1", "table1", rows, csvFiles, "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("table1", result.get(TABLE_NAME));
        assertEquals(rows, result.get(ROWS));
        assertEquals(csvFiles, result.get(ROWS_FROM_CSV_FILE));
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testAddRowsToSheetMetaToMap() {
        List<String> rows = Arrays.asList("row1");
        Map<String, Object> result = StateMapperUtil.addRowsToSheetMetaToMap("drive1", "file1", "sheet1", rows, null, "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("sheet1", result.get(SHEET_NAME));
        assertEquals(rows, result.get(ROWS));
        assertNull(result.get(ROWS_FROM_CSV_FILE));
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testRangeSumMetaToMap() {
        Map<String, Object> result = StateMapperUtil.rangeSumMetaToMap("drive1", "file1", "sheet1", "A1:B2", "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("sheet1", result.get(SHEET_NAME));
        assertEquals("A1:B2", result.get(RANGE));
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testGetTableRowsMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getTableRowsMetaToMap("drive1", "file1", "table1", "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("table1", result.get(TABLE_NAME));
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testSheetVlookupMetaToMap() {
        Map<String, Object> result = StateMapperUtil.sheetVlookupMetaToMap("drive1", "file1", "value1",
                "sheet1", "A1:C10", 2.0, true, "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("value1", result.get(LOOKUP_VALUE));
        assertEquals("sheet1", result.get(SHEET_NAME));
        assertEquals("A1:C10", result.get(RANGE));
        assertEquals(2.0, result.get(COLUMN_INDEX));
        assertEquals(true, result.get(EXACT_MATCH));
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testAddDataToCellMetaToMap() {
        Map<String, Object> result = StateMapperUtil.addDataToCellMetaToMap("drive1", "file1", "sheet1", "A1", "data1", "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("A1", result.get(CELL_ID));
        assertEquals("data1", result.get(DATA));
    }

    @Test
    void testGetRowsForRangeMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getRowsForRangeMetaToMap("drive1", "file1", "sheet1", "A1:B5", "state1");

        assertNotNull(result);
        assertEquals("A1:B5", result.get(RANGE));
        assertEquals("sheet1", result.get(SHEET_NAME));
    }

    @Test
    void testVlookupInTableMetaToMap() {
        Map<String, Object> result = StateMapperUtil.vlookupInTableMetaToMap("drive1", "file1", "value1",
                "table1", 3.0, false, "state1");

        assertNotNull(result);
        assertEquals("table1", result.get(TABLE_NAME));
        assertEquals("value1", result.get(LOOKUP_VALUE));
        assertEquals(3.0, result.get(COLUMN_INDEX));
        assertEquals(false, result.get(EXACT_MATCH));
    }

    @Test
    void testDeleteRowWithIndexMetaToMap() {
        Map<String, Object> result = StateMapperUtil.deleteRowWithIndexMetaToMap("drive1", "file1", "sheet1", 5.0, "state1");

        assertNotNull(result);
        assertEquals(5.0, result.get(ROW_INDEX));
    }

    @Test
    void testGetCompleteRowInfoMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getCompleteRowInfoMetaToMap("drive1", "file1", "sheet1", 3.0, "state1");

        assertNotNull(result);
        assertEquals(3.0, result.get(ROW_INDEX));
    }

    @Test
    void testDeleteColumnMetaToMap() {
        Map<String, Object> result = StateMapperUtil.deleteColumnMetaToMap("drive1", "file1", "table1", "col1", "state1");

        assertNotNull(result);
        assertEquals("col1", result.get(COLUMN_NAME));
        assertEquals("table1", result.get(TABLE_NAME));
    }

    @Test
    void testExecuteFormulaMetaToMap() {
        Map<String, Object> result = StateMapperUtil.executeFormulaMetaToMap("drive1", "file1", "=SUM(A1:A5)", "SUM", "total", "state1");

        assertNotNull(result);
        assertEquals("=SUM(A1:A5)", result.get(FORMULA));
        assertEquals("SUM", result.get(FORMULA_NAME));
        assertEquals("total", result.get(COMMENT));
    }

    @Test
    void testSearchMetaToMap() {
        Map<String, Object> result = StateMapperUtil.searchMetaToMap("drive1", "file1", "sheet1", "searchTerm", "state1");

        assertNotNull(result);
        assertEquals("searchTerm", result.get(DATA_TO_SEARCH));
    }

    @Test
    void testCellDataMetaToMap() {
        Map<String, Object> result = StateMapperUtil.cellDataMetaToMap("drive1", "file1", "sheet1", "B2", "state1");

        assertNotNull(result);
        assertEquals("B2", result.get(CELL_ID));
    }

    @Test
    void testGetSharedExcelFileMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getSharedExcelFileMetaToMap("folder1", "file1.xlsx", "state1");

        assertNotNull(result);
        assertEquals("folder1", result.get(SHARED_FOLDER_NAME));
        assertEquals("file1.xlsx", result.get(SHARED_FILE_NAME));
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testGetColumnValuesForKeysMetaToMap() {
        List<String> keyRows = Arrays.asList("key1", "key2");
        List<String> lookupColumn = Arrays.asList("col1");
        Map<String, Object> result = StateMapperUtil.getColumnValuesForKeysMetaToMap("drive1", "file1", "table1",
                keyRows, lookupColumn, true, "state1");

        assertNotNull(result);
        assertEquals(keyRows, result.get(KEY_ROWS));
        assertEquals(lookupColumn, result.get(LOOKUP_COLUMN));
        assertEquals(true, result.get(GET_ALL_RECORD));
    }

    @Test
    void testColumnsSumMetaToMap() {
        List<String> columns = Arrays.asList("col1", "col2");
        Map<String, Object> result = StateMapperUtil.columnsSumMetaToMap("drive1", "file1", "table1", columns, "state1");

        assertNotNull(result);
        assertEquals(columns, result.get(COLUMNS));
    }

    @Test
    void testConvertToPdfMetaToMap() {
        Map<String, Object> result = StateMapperUtil.convertToPdfMetaToMap("drive1", "file1", "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testDeleteColumnWithIndexMetaToMap() {
        Map<String, Object> result = StateMapperUtil.deleteColumnWithIndexMetaToMap("drive1", "file1", "sheet1", "col1", "state1");

        assertNotNull(result);
        assertEquals("col1", result.get(COLUMN_NAME));
        assertEquals("sheet1", result.get(SHEET_NAME));
    }

    @Test
    void testAddRowWithIndexMetaToMapWithDriveId() {
        File mockFile = mock(File.class);
        List<String> rows = Arrays.asList("row1");
        List<File> csvFiles = Collections.singletonList(mockFile);

        Map<String, Object> result = StateMapperUtil.addRowWithIndexMetaToMap("drive1", "file1", "sheet1",
                rows, csvFiles, 2.0, "state1");

        assertNotNull(result);
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals(2.0, result.get(ROW_INDEX));
    }

    @Test
    void testUpdateRowWithIndexMetaToMap() {
        Map<String, Object> result = StateMapperUtil.updateRowWithIndexMetaToMap("drive1", "file1", "sheet1",
                "rowData", 3.0, "state1");

        assertNotNull(result);
        assertEquals("rowData", result.get(ROW));
        assertEquals(3.0, result.get(ROW_INDEX));
    }

    @Test
    void testGetColumnValuesForKeysInSheetMetaToMap() {
        List<String> lookupRows = Arrays.asList("r1", "r2");
        List<String> lookupColumns = Arrays.asList("c1");

        Map<String, Object> result = StateMapperUtil.getColumnValuesForKeysInSheetMetaToMap("drive1", "file1", "sheet1",
                "A1:C10", lookupRows, lookupColumns, true, 2.0, "state1");

        assertNotNull(result);
        assertEquals(lookupRows, result.get(LOOKUP_ROWS));
        assertEquals(lookupColumns, result.get(LOOKUP_COLUMNS));
        assertEquals(2.0, result.get(COLUMN_NUMBER));
    }

    @Test
    void testGetColumnValuesForKeysInSheetMetaToMapWithNullColumnNumber() {
        Map<String, Object> result = StateMapperUtil.getColumnValuesForKeysInSheetMetaToMap("drive1", "file1", "sheet1",
                "A1:C10", null, null, false, null, "state1");

        assertNotNull(result);
        assertFalse(result.containsKey(COLUMN_NUMBER));
    }

    @Test
    void testAddMultipleCellDataMetaToMap() {
        List<Map<String, Object>> cellData = List.of(Map.of("cell", "A1", "value", "test"));
        Map<String, Object> result = StateMapperUtil.addMultipleCellDataMetaToMap("drive1", "file1", "sheet1",
                cellData, "state1");

        assertNotNull(result);
        assertEquals(cellData, result.get(CELL_ID_DATA));
    }

    @Test
    void testAddMultipleCellDataMetaToMapForSheet() {
        List<Map<String, Object>> cellData = List.of(Map.of("cell", "A1", "value", "test"));
        Map<String, Object> result = StateMapperUtil.addMultipleCellDataMetaToMapForSheet("file1", "sheet1",
                cellData, "state1");

        assertNotNull(result);
        assertFalse(result.containsKey(DRIVE_ID));
        assertEquals(cellData, result.get(CELL_ID_DATA));
    }

    @Test
    void testClearSharedSheetMetaToMap() {
        Map<String, Object> result = StateMapperUtil.clearSharedSheetMetaToMap("drive1", "file1", "sheet1", "A1:Z100", "state1");

        assertNotNull(result);
        assertEquals("A1:Z100", result.get(RANGE));
    }

    @Test
    void testDeleteSheetMetaToMap() {
        Map<String, Object> result = StateMapperUtil.deleteSheetMetaToMap("file1", "sheet1", "state1");

        assertNotNull(result);
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("sheet1", result.get(SHEET_NAME));
        assertFalse(result.containsKey(DRIVE_ID));
    }

    @Test
    void testGetCellValueMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getCellValueMetaToMap("file1", "sheet1", "A1", "state1");

        assertNotNull(result);
        assertEquals("A1", result.get(CELL_ID));
    }

    @Test
    void testAddRowsMetaToMap() {
        List<String> rows = Arrays.asList("r1", "r2");
        Map<String, Object> result = StateMapperUtil.addRowsMetaToMap("file1", "sheet1", rows, null, "state1");

        assertNotNull(result);
        assertEquals(rows, result.get(ROWS));
    }

    @Test
    void testAddRowWithIndexMetaToMapWithoutDriveId() {
        List<String> rows = Arrays.asList("r1");
        Map<String, Object> result = StateMapperUtil.addRowWithIndexMetaToMap("file1", "sheet1", rows, null, 1.0, "state1");

        assertNotNull(result);
        assertEquals(1.0, result.get(ROW_INDEX));
        assertFalse(result.containsKey(DRIVE_ID));
    }

    @Test
    void testFindValueMetaToMap() {
        Map<String, Object> result = StateMapperUtil.findValueMetaToMap("file1", "sheet1", "lookupVal", "state1");

        assertNotNull(result);
        assertEquals("lookupVal", result.get(LOOKUP_VALUE));
    }

    @Test
    void testGetExcelFileInfoMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getExcelFileInfoMetaToMap("/path/to/file", "file.xlsx", "state1");

        assertNotNull(result);
        assertEquals("/path/to/file", result.get(PATH));
        assertEquals("file.xlsx", result.get(NAME));
    }

    @Test
    void testGetSheetsInfoMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getSheetsInfoMetaToMap("file1", "state1");

        assertNotNull(result);
        assertEquals("file1", result.get(FILE_ID));
    }

    @Test
    void testCreateSheetMetaToMap() {
        Map<String, Object> result = StateMapperUtil.createSheetMetaToMap("file1", "newSheet", "state1");

        assertNotNull(result);
        assertEquals("newSheet", result.get(SHEET_NAME));
    }

    @Test
    void testClearSheetForRangeMetaToMap() {
        Map<String, Object> result = StateMapperUtil.clearSheetForRangeMetaToMap("file1", "sheet1", "A1:Z10", "state1");

        assertNotNull(result);
        assertEquals("A1:Z10", result.get(RANGE));
    }

    @Test
    void testVlookupMetaToMap() {
        Map<String, Object> result = StateMapperUtil.vlookupMetaToMap("file1", "val", "sheet1", "A1:C5", 2.0, true, "state1");

        assertNotNull(result);
        assertEquals(true, result.get(EXACT_MATCH));
    }

    @Test
    void testVlookupMetaToMapWithNullExactMatch() {
        Map<String, Object> result = StateMapperUtil.vlookupMetaToMap("file1", "val", "sheet1", "A1:C5", 2.0, null, "state1");

        assertNotNull(result);
        assertEquals(false, result.get(EXACT_MATCH));
    }

    @Test
    void testAddFormulaToCellMetaToMap() {
        Map<String, Object> result = StateMapperUtil.addFormulaToCellMetaToMap("file1", "sheet1", "A1", "=SUM(B1:B5)", "state1");

        assertNotNull(result);
        assertEquals("=SUM(B1:B5)", result.get(FORMULA));
        assertEquals("A1", result.get(CELL_ID));
    }

    @Test
    void testSummationMetaToMap() {
        Map<String, Object> result = StateMapperUtil.summationMetaToMap("file1", "sheet1", "A1:A10", "state1");

        assertNotNull(result);
        assertEquals("A1:A10", result.get(RANGE));
    }

    @Test
    void testUploadFileAndConvertToPDFMetaToMap() {
        File mockFile = mock(File.class);
        Map<String, Object> result = StateMapperUtil.uploadFileAndConvertToPDFMetaToMap(mockFile, "/path/file.xlsx", "state1");

        assertNotNull(result);
        assertEquals(mockFile, result.get(EXCEL_FILE));
        assertEquals("/path/file.xlsx", result.get(FILE_PATH));
    }

    @Test
    void testGetMinMaxValueMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getMinMaxValueMetaToMap("file1", "sheet1", "KeyCol", "A:B", "state1");

        assertNotNull(result);
        assertEquals("KeyCol", result.get(KEY_COLUMN_NAME));
        assertEquals("A:B", result.get(COLUMN_RANGE));
    }

    @Test
    void testVlookupTableMetaToMap() {
        Map<String, Object> result = StateMapperUtil.vlookupTableMetaToMap("file1", "val", "table1", 1.0, true, "state1");

        assertNotNull(result);
        assertEquals(true, result.get(EXACT_MATCH));
    }

    @Test
    void testVlookupTableMetaToMapWithNullExactMatch() {
        Map<String, Object> result = StateMapperUtil.vlookupTableMetaToMap("file1", "val", "table1", 1.0, null, "state1");

        assertNotNull(result);
        assertEquals(false, result.get(EXACT_MATCH));
    }

    @Test
    void testAddTableRowsMetaToMap() {
        List<String> rows = Arrays.asList("r1");
        Map<String, Object> result = StateMapperUtil.addTableRowsMetaToMap("file1", "table1", rows, null, "state1");

        assertNotNull(result);
        assertEquals(rows, result.get(ROWS));
        assertEquals("table1", result.get(TABLE_NAME));
    }

    @Test
    void testAddRowsToTableWithIndexMetaToMap() {
        List<String> rows = Arrays.asList("r1");
        Map<String, Object> result = StateMapperUtil.addRowsToTableWithIndexMetaToMap("drive1", "file1", "table1",
                rows, null, 5.0, ",", "state1");

        assertNotNull(result);
        assertEquals(",", result.get(SEPARATOR));
        assertEquals(5.0, result.get(ROW_INDEX));
    }

    @Test
    void testGetRowsMetaToMap() {
        Map<String, Object> result = StateMapperUtil.getRowsMetaToMap("workbook1", "table1", "state1");

        assertNotNull(result);
        assertEquals("workbook1", result.get(FILE_ID));
        assertEquals("table1", result.get(TABLE_NAME));
    }

    @Test
    void testTableSummationMetaToMap() {
        List<String> columns = Arrays.asList("c1", "c2");
        Map<String, Object> result = StateMapperUtil.tableSummationMetaToMap("file1", "table1", columns, "state1");

        assertNotNull(result);
        assertEquals(columns, result.get(COLUMNS));
    }

    @Test
    void testDeleteTableColumnMetaToMap() {
        Map<String, Object> result = StateMapperUtil.deleteTableColumnMetaToMap("file1", "table1", "col1", "state1");

        assertNotNull(result);
        assertEquals("col1", result.get(COLUMN_NAME));
    }

    @Test
    void testCreateTableForRangeMetaToMap() {
        Map<String, Object> result = StateMapperUtil.createTableForRangeMetaToMap("file1", "sheet1", "A1:D10", "drive1", "state1");

        assertNotNull(result);
        assertEquals("A1:D10", result.get(RANGE));
        assertEquals("drive1", result.get(DRIVE_ID));
    }

    @Test
    void testGetTableColumnValuesForKeysMetaToMap() {
        List<String> keyRows = Arrays.asList("k1");
        List<String> lookupCols = Arrays.asList("c1");
        Map<String, Object> result = StateMapperUtil.getTableColumnValuesForKeysMetaToMap("file1", "table1",
                keyRows, lookupCols, true, 1.0, "state1");

        assertNotNull(result);
        assertEquals(keyRows, result.get(KEY_ROWS));
        assertEquals(lookupCols, result.get(LOOKUP_COLUMNS));
        assertEquals(1.0, result.get(COLUMN_NUMBER));
    }

    @Test
    void testPrivateConstructor() throws Exception {
        Constructor<StateMapperUtil> constructor = StateMapperUtil.class.getDeclaredConstructor();
        assertTrue(java.lang.reflect.Modifier.isPrivate(constructor.getModifiers()));
        constructor.setAccessible(true);
        assertNotNull(constructor.newInstance());
    }
}
