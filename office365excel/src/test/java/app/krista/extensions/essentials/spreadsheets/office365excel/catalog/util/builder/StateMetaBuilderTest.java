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
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.mock;

class StateMetaBuilderTest {

    @Test
    void testEmptyBuilderReturnsEmptyMap() {
        Map<String, Object> result = StateMetaBuilder.builder().build();
        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void testBuilderFactory() {
        StateMetaBuilder builder = StateMetaBuilder.builder();
        assertNotNull(builder);
    }

    @Test
    void testWithDriveId() {
        Map<String, Object> result = StateMetaBuilder.builder().withDriveId("drive1").build();
        assertEquals("drive1", result.get(DRIVE_ID));
    }

    @Test
    void testWithFileId() {
        Map<String, Object> result = StateMetaBuilder.builder().withFileId("file1").build();
        assertEquals("file1", result.get(FILE_ID));
    }

    @Test
    void testWithTableName() {
        Map<String, Object> result = StateMetaBuilder.builder().withTableName("table1").build();
        assertEquals("table1", result.get(TABLE_NAME));
    }

    @Test
    void testWithSheetName() {
        Map<String, Object> result = StateMetaBuilder.builder().withSheetName("Sheet1").build();
        assertEquals("Sheet1", result.get(SHEET_NAME));
    }

    @Test
    void testWithRows() {
        List<String> rows = Arrays.asList("r1", "r2");
        Map<String, Object> result = StateMetaBuilder.builder().withRows(rows).build();
        assertEquals(rows, result.get(ROWS));
    }

    @Test
    void testWithRowsFromCsvFile() {
        File mockFile = mock(File.class);
        List<File> csvFiles = List.of(mockFile);
        Map<String, Object> result = StateMetaBuilder.builder().withRowsFromCsvFile(csvFiles).build();
        assertEquals(csvFiles, result.get(ROWS_FROM_CSV_FILE));
    }

    @Test
    void testWithStateId() {
        Map<String, Object> result = StateMetaBuilder.builder().withStateId("state1").build();
        assertEquals("state1", result.get(STATE_ID));
    }

    @Test
    void testWithRange() {
        Map<String, Object> result = StateMetaBuilder.builder().withRange("A1:B5").build();
        assertEquals("A1:B5", result.get(RANGE));
    }

    @Test
    void testWithLookupValue() {
        Map<String, Object> result = StateMetaBuilder.builder().withLookupValue("val").build();
        assertEquals("val", result.get(LOOKUP_VALUE));
    }

    @Test
    void testWithColumnIndex() {
        Map<String, Object> result = StateMetaBuilder.builder().withColumnIndex(3.0).build();
        assertEquals(3.0, result.get(COLUMN_INDEX));
    }

    @Test
    void testWithExactMatch() {
        Map<String, Object> result = StateMetaBuilder.builder().withExactMatch(true).build();
        assertEquals(true, result.get(EXACT_MATCH));
    }

    @Test
    void testWithCellId() {
        Map<String, Object> result = StateMetaBuilder.builder().withCellId("A1").build();
        assertEquals("A1", result.get(CELL_ID));
    }

    @Test
    void testWithData() {
        Map<String, Object> result = StateMetaBuilder.builder().withData("data1").build();
        assertEquals("data1", result.get(DATA));
    }

    @Test
    void testWithRowIndex() {
        Map<String, Object> result = StateMetaBuilder.builder().withRowIndex(5.0).build();
        assertEquals(5.0, result.get(ROW_INDEX));
    }

    @Test
    void testWithColumnName() {
        Map<String, Object> result = StateMetaBuilder.builder().withColumnName("col1").build();
        assertEquals("col1", result.get(COLUMN_NAME));
    }

    @Test
    void testWithFormula() {
        Map<String, Object> result = StateMetaBuilder.builder().withFormula("=SUM(A1:A5)").build();
        assertEquals("=SUM(A1:A5)", result.get(FORMULA));
    }

    @Test
    void testWithFormulaName() {
        Map<String, Object> result = StateMetaBuilder.builder().withFormulaName("SUM").build();
        assertEquals("SUM", result.get(FORMULA_NAME));
    }

    @Test
    void testWithComment() {
        Map<String, Object> result = StateMetaBuilder.builder().withComment("comment").build();
        assertEquals("comment", result.get(COMMENT));
    }

    @Test
    void testWithDataToSearch() {
        Map<String, Object> result = StateMetaBuilder.builder().withDataToSearch("search").build();
        assertEquals("search", result.get(DATA_TO_SEARCH));
    }

    @Test
    void testWithKeyRows() {
        List<String> keyRows = Arrays.asList("k1", "k2");
        Map<String, Object> result = StateMetaBuilder.builder().withKeyRows(keyRows).build();
        assertEquals(keyRows, result.get(KEY_ROWS));
    }

    @Test
    void testWithLookupColumn() {
        List<String> lookupCol = Arrays.asList("c1");
        Map<String, Object> result = StateMetaBuilder.builder().withLookupColumn(lookupCol).build();
        assertEquals(lookupCol, result.get(LOOKUP_COLUMN));
    }

    @Test
    void testWithGetAllRecord() {
        Map<String, Object> result = StateMetaBuilder.builder().withGetAllRecord(true).build();
        assertEquals(true, result.get(GET_ALL_RECORD));
    }

    @Test
    void testWithSharedFolderName() {
        Map<String, Object> result = StateMetaBuilder.builder().withSharedFolderName("folder").build();
        assertEquals("folder", result.get(SHARED_FOLDER_NAME));
    }

    @Test
    void testWithSharedFileName() {
        Map<String, Object> result = StateMetaBuilder.builder().withSharedFileName("file.xlsx").build();
        assertEquals("file.xlsx", result.get(SHARED_FILE_NAME));
    }

    @Test
    void testWithPath() {
        Map<String, Object> result = StateMetaBuilder.builder().withPath("/path").build();
        assertEquals("/path", result.get(PATH));
    }

    @Test
    void testWithName() {
        Map<String, Object> result = StateMetaBuilder.builder().withName("name").build();
        assertEquals("name", result.get(NAME));
    }

    @Test
    void testWithExcelFile() {
        File mockFile = mock(File.class);
        Map<String, Object> result = StateMetaBuilder.builder().withExcelFile(mockFile).build();
        assertEquals(mockFile, result.get(EXCEL_FILE));
    }

    @Test
    void testWithFilePath() {
        Map<String, Object> result = StateMetaBuilder.builder().withFilePath("/path/file.xlsx").build();
        assertEquals("/path/file.xlsx", result.get(FILE_PATH));
    }

    @Test
    void testWithKeyColumnName() {
        Map<String, Object> result = StateMetaBuilder.builder().withKeyColumnName("keyCol").build();
        assertEquals("keyCol", result.get(KEY_COLUMN_NAME));
    }

    @Test
    void testWithColumnRange() {
        Map<String, Object> result = StateMetaBuilder.builder().withColumnRange("A:B").build();
        assertEquals("A:B", result.get(COLUMN_RANGE));
    }

    @Test
    void testWithSeparator() {
        Map<String, Object> result = StateMetaBuilder.builder().withSeparator(",").build();
        assertEquals(",", result.get(SEPARATOR));
    }

    @Test
    void testWithColumns() {
        List<String> columns = Arrays.asList("c1", "c2");
        Map<String, Object> result = StateMetaBuilder.builder().withColumns(columns).build();
        assertEquals(columns, result.get(COLUMNS));
    }

    @Test
    void testWithColumnNumber() {
        Map<String, Object> result = StateMetaBuilder.builder().withColumnNumber(4.0).build();
        assertEquals(4.0, result.get(COLUMN_NUMBER));
    }

    @Test
    void testWithAdditional() {
        Map<String, Object> result = StateMetaBuilder.builder().withAdditional("customKey", "customVal").build();
        assertEquals("customVal", result.get("customKey"));
    }

    @Test
    void testWithLookupRows() {
        List<String> lookupRows = Arrays.asList("lr1", "lr2");
        Map<String, Object> result = StateMetaBuilder.builder().withLookupRows(lookupRows).build();
        assertEquals(lookupRows, result.get(LOOKUP_ROWS));
    }

    @Test
    void testWithLookupColumns() {
        List<String> lookupCols = Arrays.asList("lc1");
        Map<String, Object> result = StateMetaBuilder.builder().withLookupColumns(lookupCols).build();
        assertEquals(lookupCols, result.get(LOOKUP_COLUMNS));
    }

    @Test
    void testChainingMultipleMethods() {
        Map<String, Object> result = StateMetaBuilder.builder()
                .withDriveId("drive1")
                .withFileId("file1")
                .withTableName("table1")
                .withSheetName("Sheet1")
                .withStateId("state1")
                .withRange("A1:B5")
                .build();

        assertEquals(6, result.size());
        assertEquals("drive1", result.get(DRIVE_ID));
        assertEquals("file1", result.get(FILE_ID));
        assertEquals("table1", result.get(TABLE_NAME));
        assertEquals("Sheet1", result.get(SHEET_NAME));
        assertEquals("state1", result.get(STATE_ID));
        assertEquals("A1:B5", result.get(RANGE));
    }

    @Test
    void testBuildReturnsSameMapInstance() {
        StateMetaBuilder builder = StateMetaBuilder.builder().withDriveId("d1");
        Map<String, Object> map1 = builder.build();
        Map<String, Object> map2 = builder.build();
        // build() returns the internal map directly
        assertSame(map1, map2);
    }

    @Test
    void testWithNullValues() {
        Map<String, Object> result = StateMetaBuilder.builder()
                .withDriveId(null)
                .withFileId(null)
                .build();

        assertTrue(result.containsKey(DRIVE_ID));
        assertNull(result.get(DRIVE_ID));
        assertTrue(result.containsKey(FILE_ID));
        assertNull(result.get(FILE_ID));
    }
}
