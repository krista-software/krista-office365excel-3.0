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
import java.util.*;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryResources.*;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.mock;

class TelemetryTagsHelperTest {

    // TelemetryTagsBuilder tests
    @Test
    void testBuilderOperationType() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .operationType("test_op")
                .build();

        assertEquals("test_op", tags.get(OPERATION_TYPE));
    }

    @Test
    void testBuilderStringTag() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .stringTag("key", "value")
                .build();

        assertEquals("true", tags.get("key"));
    }

    @Test
    void testBuilderStringTagWithNull() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .stringTag("key", null)
                .build();

        assertEquals("false", tags.get("key"));
    }

    @Test
    void testBuilderStringTagWithEmpty() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .stringTag("key", "")
                .build();

        assertEquals("false", tags.get("key"));
    }

    @Test
    void testBuilderStringTagWithWhitespace() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .stringTag("key", "   ")
                .build();

        assertEquals("false", tags.get("key"));
    }

    @Test
    void testBuilderListTag() {
        List<String> list = Arrays.asList("a", "b", "c");
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .listTag("hasKey", "countKey", list)
                .build();

        assertEquals("true", tags.get("hasKey"));
        assertEquals("3", tags.get("countKey"));
    }

    @Test
    void testBuilderListTagWithNull() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .listTag("hasKey", "countKey", null)
                .build();

        assertEquals("false", tags.get("hasKey"));
        assertEquals("0", tags.get("countKey"));
    }

    @Test
    void testBuilderListTagWithEmptyList() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .listTag("hasKey", "countKey", Collections.emptyList())
                .build();

        assertEquals("false", tags.get("hasKey"));
        assertEquals("0", tags.get("countKey"));
    }

    @Test
    void testBuilderRowsTag() {
        List<String> rows = Arrays.asList("row1", "row2");
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .rowsTag(rows)
                .build();

        assertEquals("true", tags.get(HAS_ROWS));
        assertEquals("2", tags.get(ROW_COUNT));
    }

    @Test
    void testBuilderRowsTagWithNullRows() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .rowsTag(null)
                .build();

        assertEquals("false", tags.get(HAS_ROWS));
        assertEquals("0", tags.get(ROW_COUNT));
    }

    @Test
    void testBuilderRowsTagWithAllNullEntries() {
        List<String> rows = Arrays.asList(null, null);
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .rowsTag(rows)
                .build();

        assertEquals("false", tags.get(HAS_ROWS));
        assertEquals("2", tags.get(ROW_COUNT));
    }

    @Test
    void testBuilderStringWithLength() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .stringWithLength("hasKey", "lengthKey", "hello")
                .build();

        assertEquals("true", tags.get("hasKey"));
        assertEquals("5", tags.get("lengthKey"));
    }

    @Test
    void testBuilderStringWithLengthNull() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .stringWithLength("hasKey", "lengthKey", null)
                .build();

        assertEquals("false", tags.get("hasKey"));
        assertEquals("0", tags.get("lengthKey"));
    }

    @Test
    void testBuilderBooleanTag() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .booleanTag("key", true)
                .build();

        assertEquals("true", tags.get("key"));
    }

    @Test
    void testBuilderBooleanTagNull() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .booleanTag("key", null)
                .build();

        assertEquals("null", tags.get("key"));
    }

    @Test
    void testBuilderNullableTag() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .nullableTag("key", "notNull")
                .build();

        assertEquals("true", tags.get("key"));
    }

    @Test
    void testBuilderNullableTagWithNull() {
        Map<String, String> tags = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .nullableTag("key", null)
                .build();

        assertEquals("false", tags.get("key"));
    }

    @Test
    void testBuilderBuildReturnsNewMap() {
        TelemetryTagsHelper.TelemetryTagsBuilder builder = new TelemetryTagsHelper.TelemetryTagsBuilder()
                .operationType("test");

        Map<String, String> tags1 = builder.build();
        Map<String, String> tags2 = builder.build();

        assertNotSame(tags1, tags2);
        assertEquals(tags1, tags2);
    }

    // Static factory method tests
    @Test
    void testCreateTableOperationTags() {
        List<String> rows = Arrays.asList("r1");
        File mockFile = mock(File.class);
        List<File> csvFiles = Collections.singletonList(mockFile);

        Map<String, String> tags = TelemetryTagsHelper.createTableOperationTags(rows, csvFiles);

        assertEquals(TABLE_INSERT, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_ROWS));
        assertEquals("true", tags.get(HAS_CSV_FILE));
    }

    @Test
    void testCreateSheetOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createSheetOperationTags(null, null);

        assertEquals(SHEET_INSERT, tags.get(OPERATION_TYPE));
        assertEquals("false", tags.get(HAS_ROWS));
        assertEquals("false", tags.get(HAS_FILE));
    }

    @Test
    void testCreateRangeOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createRangeOperationTags("A1:B2", "Sheet1");

        assertEquals(RANGE_SUM, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_RANGE));
        assertEquals("true", tags.get(SHEET_PROVIDED));
    }

    @Test
    void testCreateVlookupOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createVlookupOperationTags("val", "A1:C5", "Sheet1");

        assertEquals(VLOOKUP, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_LOOKUP_VALUE));
    }

    @Test
    void testCreateTableQueryOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createTableQueryOperationTags("Table1");

        assertEquals(TABLE_QUERY, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_TABLE_NAME));
    }

    @Test
    void testCreateCellOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createCellOperationTags("A1", "data");

        assertEquals(CELL_UPDATE, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_CELL_ID));
        assertEquals("true", tags.get(HAS_DATA));
        assertEquals("4", tags.get(DATA_LENGTH));
    }

    @Test
    void testCreateDeleteColumnOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createDeleteColumnOperationTags("table1", "col1");

        assertEquals(DELETE_COLUMN, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateColumnValuesForKeysOperationTags() {
        List<String> keyRows = Arrays.asList("k1");
        List<String> lookupCol = Arrays.asList("c1", "c2");

        Map<String, String> tags = TelemetryTagsHelper.createColumnValuesForKeysOperationTags("table1", keyRows, lookupCol);

        assertEquals(COLUMN_VALUES_FOR_KEYS, tags.get(OPERATION_TYPE));
        assertEquals("1", tags.get(KEY_ROWS_COUNT));
        assertEquals("2", tags.get(LOOKUP_COLUMN_COUNT));
    }

    @Test
    void testCreateExecuteFormulaOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createExecuteFormulaOperationTags("=SUM(A1:A5)", "SUM");

        assertEquals(EXECUTE_FORMULA, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_FORMULA));
        assertEquals("11", tags.get(FORMULA_LENGTH));
    }

    @Test
    void testCreateVlookupInTableOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createVlookupInTableOperationTags("val", "table1", 2.0);

        assertEquals(VLOOKUP_IN_TABLE, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_COLUMN_INDEX));
    }

    @Test
    void testCreateVlookupInTableOperationTagsWithNullColumnIndex() {
        Map<String, String> tags = TelemetryTagsHelper.createVlookupInTableOperationTags("val", "table1", null);

        assertEquals("false", tags.get(HAS_COLUMN_INDEX));
    }

    @Test
    void testCreateColumnsSumOperationTags() {
        List<String> columns = Arrays.asList("c1", "c2");
        Map<String, String> tags = TelemetryTagsHelper.createColumnsSumOperationTags("table1", columns);

        assertEquals(COLUMNS_SUM, tags.get(OPERATION_TYPE));
        assertEquals("2", tags.get(COLUMN_COUNT));
    }

    @Test
    void testCreateSearchOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createSearchOperationTags("Sheet1", "search");

        assertEquals(SEARCH, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateCellDataOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createCellDataOperationTags("Sheet1", "A1");

        assertEquals(CELL_DATA, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateGetSharedExcelFileOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createGetSharedExcelFileOperationTags("folder", "file.xlsx");

        assertEquals(GET_SHARED_EXCEL_FILE, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateConvertToPdfOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createConvertToPdfOperationTags("drive1", "file1");

        assertEquals(CONVERT_TO_PDF, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateDeleteColumnWithIndexOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createDeleteColumnWithIndexOperationTags("Sheet1", "col1");

        assertEquals(DELETE_COLUMN_WITH_INDEX, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateAddRowWithIndexOperationTags() {
        List<String> rows = Arrays.asList("r1");
        Map<String, String> tags = TelemetryTagsHelper.createAddRowWithIndexOperationTags("Sheet1", rows, null, 1.0);

        assertEquals(ADD_ROW_WITH_INDEX, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_ROW_INDEX));
    }

    @Test
    void testCreateUpdateRowWithIndexOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createUpdateRowWithIndexOperationTags("Sheet1", "rowData", 2.0);

        assertEquals(UPDATE_ROW_WITH_INDEX, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateGetCompleteRowInfoOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createGetCompleteRowInfoOperationTags("Sheet1", 3.0);

        assertEquals(GET_COMPLETE_ROW_INFO, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateDeleteRowWithIndexOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createDeleteRowWithIndexOperationTags("Sheet1", 1.0);

        assertEquals(DELETE_ROW_WITH_INDEX, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateAddMultipleCellDataOperationTags() {
        List<Map<String, Object>> cellData = List.of(Map.of("cell", "A1"));
        Map<String, String> tags = TelemetryTagsHelper.createAddMultipleCellDataOperationTags("Sheet1", cellData);

        assertEquals(ADD_MULTIPLE_CELL_DATA, tags.get(OPERATION_TYPE));
        assertEquals("1", tags.get(CELL_DATA_COUNT));
    }

    @Test
    void testCreateGetColumnValuesForKeysInSheetOperationTagsFull() {
        List<String> lookupRows = Arrays.asList("r1");
        List<String> lookupCols = Arrays.asList("c1");
        Map<String, String> tags = TelemetryTagsHelper.createGetColumnValuesForKeysInSheetOperationTags(
                "Sheet1", "A1:B5", lookupRows, lookupCols, true, 1.0);

        assertEquals(GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(GET_ALL_RECORDS));
        assertEquals("true", tags.get(HAS_COLUMN_NUMBER));
    }

    @Test
    void testCreateClearSharedSheetForRangeOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createClearSharedSheetForRangeOperationTags("Sheet1", "A1:Z100");

        assertEquals(CLEAR_SHARED_SHEET_FOR_RANGE, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateDeleteSheetOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createDeleteSheetOperationTags("Sheet1");

        assertEquals(DELETE_SHEET, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateGetCellValueOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createGetCellValueOperationTags("file1", "Sheet1", "A1");

        assertEquals(GET_CELL_VALUE, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateAddRowsOperationTags() {
        List<String> rows = Arrays.asList("r1");
        Map<String, String> tags = TelemetryTagsHelper.createAddRowsOperationTags(rows, null);

        assertEquals(ADD_ROWS, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateFindValueOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createFindValueOperationTags("Sheet1", "val");

        assertEquals(FIND_VALUE, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateGetExcelFileInfoOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createGetExcelFileInfoOperationTags("/path", "file.xlsx");

        assertEquals(GET_EXCEL_FILE_INFO, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateGetSheetsInfoOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createGetSheetsInfoOperationTags("file1");

        assertEquals(GET_SHEETS_INFO, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateAddDataToCellOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createAddDataToCellOperationTags("A1", "data");

        assertEquals(ADD_DATA_TO_CELL, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateCreateSheetOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createCreateSheetOperationTags("NewSheet");

        assertEquals(CREATE_SHEET, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateClearSheetForRangeOperationTagsForSheet() {
        Map<String, String> tags = TelemetryTagsHelper.createClearSheetForRangeOperationTagsForSheet("Sheet1", "A1:Z100");

        assertEquals(CLEAR_SHEET_FOR_RANGE, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateAddFormulaToCellOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createAddFormulaToCellOperationTags("A1", "=SUM(B1:B5)");

        assertEquals(ADD_FORMULA_TO_CELL, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateMinMaxValueOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createMinMaxValueOperationTags("KeyCol", "A:B");

        assertEquals(GET_MIN_MAX_VALUE, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_KEY_COLUMN_NAME));
        assertEquals("true", tags.get(HAS_COLUMN_RANGE));
    }

    @Test
    void testCreateMinMaxValueOperationTagsWithNulls() {
        Map<String, String> tags = TelemetryTagsHelper.createMinMaxValueOperationTags(null, null);

        assertEquals("false", tags.get(HAS_KEY_COLUMN_NAME));
        assertEquals("false", tags.get(HAS_COLUMN_RANGE));
    }

    @Test
    void testCreateMinMaxValueOperationTagsWithEmpty() {
        Map<String, String> tags = TelemetryTagsHelper.createMinMaxValueOperationTags("", "  ");

        assertEquals("false", tags.get(HAS_KEY_COLUMN_NAME));
        assertEquals("false", tags.get(HAS_COLUMN_RANGE));
    }

    @Test
    void testCreateUploadFileAndConvertToPDFOperationTagsWithFile() {
        File mockFile = mock(File.class);
        Map<String, String> tags = TelemetryTagsHelper.createUploadFileAndConvertToPDFOperationTags(mockFile, "/path/file.xlsx");

        assertEquals(UPLOAD_FILE_AND_CONVERT_TO_PDF, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_EXCEL_FILE));
        assertEquals("true", tags.get(HAS_FILE_PATH));
    }

    @Test
    void testCreateUploadFileAndConvertToPDFOperationTagsWithNulls() {
        Map<String, String> tags = TelemetryTagsHelper.createUploadFileAndConvertToPDFOperationTags(null, null);

        assertEquals("false", tags.get(HAS_EXCEL_FILE));
        assertEquals("false", tags.get(HAS_FILE_PATH));
    }

    @Test
    void testCreateUploadFileAndConvertToPDFOperationTagsWithEmptyFilePath() {
        Map<String, String> tags = TelemetryTagsHelper.createUploadFileAndConvertToPDFOperationTags(null, "  ");

        assertEquals("false", tags.get(HAS_FILE_PATH));
    }

    @Test
    void testCreateTableVlookupOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createTableVlookupOperationTags("val", "table1", 1.0);

        assertEquals(TABLE_VLOOKUP, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateTableCreationOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createTableCreationOperationTags("Sheet1", "A1:D10");

        assertEquals(TABLE_CREATION, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateTableWithIndexOperationTags() {
        List<String> rows = Arrays.asList("r1");
        Map<String, String> tags = TelemetryTagsHelper.createTableWithIndexOperationTags("table1", rows, null, 1.0);

        assertEquals(ADD_ROWS_TO_TABLE_WITH_INDEX, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_INDEX));
    }

    @Test
    void testCreateTableSummationOperationTags() {
        List<String> columns = Arrays.asList("c1");
        Map<String, String> tags = TelemetryTagsHelper.createTableSummationOperationTags("table1", columns);

        assertEquals(TABLE_SUMMATION, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateAddRowsInTableOperationTags() {
        List<String> rows = Arrays.asList("r1");
        Map<String, String> tags = TelemetryTagsHelper.createAddRowsInTableOperationTags("table1", rows, null);

        assertEquals(ADD_ROWS_IN_TABLE, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateDeleteTableColumnOperationTags() {
        Map<String, String> tags = TelemetryTagsHelper.createDeleteTableColumnOperationTags("table1", "col1");

        assertEquals(DELETE_TABLE_COLUMN, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateGetColumnValuesForKeysInSheetOperationTagsSimple() {
        List<String> lookupRows = Arrays.asList("r1");
        List<String> lookupCols = Arrays.asList("c1");
        Map<String, String> tags = TelemetryTagsHelper.createGetColumnValuesForKeysInSheetOperationTags("Sheet1", lookupRows, lookupCols);

        assertEquals(GET_COLUMN_VALUES_FOR_KEYS_IN_SHEET, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateGetColumnValuesForKeysOperationTags() {
        List<String> keyRows = Arrays.asList("k1");
        List<String> lookupCols = Arrays.asList("c1");
        Map<String, String> tags = TelemetryTagsHelper.createGetColumnValuesForKeysOperationTags("table1", keyRows, lookupCols);

        assertEquals(GET_COLUMN_VALUES_FOR_KEYS_OPERATION, tags.get(OPERATION_TYPE));
    }

    @Test
    void testCreateUploadFileAndConvertToPDFOperationTagsStringOverload() {
        Map<String, String> tags = TelemetryTagsHelper.createUploadFileAndConvertToPDFOperationTags("/path/file.xlsx");

        assertEquals(UPLOAD_FILE_AND_CONVERT_TO_PDF_OPERATION, tags.get(OPERATION_TYPE));
        assertEquals("true", tags.get(HAS_FILE_PATH));
    }

    @Test
    void testPrivateConstructor() throws Exception {
        Constructor<TelemetryTagsHelper> constructor = TelemetryTagsHelper.class.getDeclaredConstructor();
        assertTrue(java.lang.reflect.Modifier.isPrivate(constructor.getModifiers()));
        constructor.setAccessible(true);
        assertNotNull(constructor.newInstance());
    }
}
