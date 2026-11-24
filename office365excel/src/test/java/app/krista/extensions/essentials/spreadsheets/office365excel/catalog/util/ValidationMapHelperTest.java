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
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.mock;

class ValidationMapHelperTest {

    // Builder tests
    @Test
    void testBuilderDriveId() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().driveId("drive1").build();
        assertEquals("drive1", result.get(Validator.ValidationResource.DRIVE_ID));
    }

    @Test
    void testBuilderDriveIdNull() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().driveId(null).build();
        assertEquals("", result.get(Validator.ValidationResource.DRIVE_ID));
    }

    @Test
    void testBuilderFileId() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().fileId("file1").build();
        assertEquals("file1", result.get(Validator.ValidationResource.FILE_ID));
    }

    @Test
    void testBuilderFileIdNull() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().fileId(null).build();
        assertEquals("", result.get(Validator.ValidationResource.FILE_ID));
    }

    @Test
    void testBuilderSheetName() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().sheetName("Sheet1").build();
        assertEquals("Sheet1", result.get(Validator.ValidationResource.SHEET_NAME));
    }

    @Test
    void testBuilderTableName() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().tableName("Table1").build();
        assertEquals("Table1", result.get(Validator.ValidationResource.TABLE_NAME));
    }

    @Test
    void testBuilderRange() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().range("A1:B5").build();
        assertEquals("A1:B5", result.get(Validator.ValidationResource.RANGE));
    }

    @Test
    void testBuilderCellId() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().cellId("A1").build();
        assertEquals("A1", result.get(Validator.ValidationResource.CELL_ID));
    }

    @Test
    void testBuilderColumnIndex() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().columnIndex(3.0).build();
        assertEquals("3.0", result.get(Validator.ValidationResource.COLUMN_INDEX));
    }

    @Test
    void testBuilderColumnIndexNull() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().columnIndex(null).build();
        assertEquals("", result.get(Validator.ValidationResource.COLUMN_INDEX));
    }

    @Test
    void testBuilderLookupValue() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().lookupValue("val").build();
        assertEquals("val", result.get(Validator.ValidationResource.LOOKUP_VALUE));
    }

    @Test
    void testBuilderRowIndex() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().rowIndex(5.0).build();
        assertEquals("5.0", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testBuilderRowIndexNull() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().rowIndex(null).build();
        assertEquals("", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testBuilderFormula() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().formula("=SUM(A1:A5)").build();
        assertEquals("=SUM(A1:A5)", result.get(Validator.ValidationResource.FORMULA));
    }

    @Test
    void testBuilderFormulaName() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().formulaName("SUM").build();
        assertEquals("SUM", result.get(Validator.ValidationResource.FORMULA_NAME));
    }

    @Test
    void testBuilderComment() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().comment("comment").build();
        assertEquals("comment", result.get(Validator.ValidationResource.COMMENT));
    }

    @Test
    void testBuilderSharedFileName() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().sharedFileName("file.xlsx").build();
        assertEquals("file.xlsx", result.get(Validator.ValidationResource.SHARED_FILE_NAME));
    }

    @Test
    void testBuilderColumnName() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().columnName("col1").build();
        assertEquals("col1", result.get(Validator.ValidationResource.COLUMN_NAME));
    }

    @Test
    void testBuilderRowData() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().rowData("data").build();
        assertEquals("data", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testBuilderRowDataNull() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder().rowData(null).build();
        assertEquals("", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testBuilderConditionalRowDataBothMissing() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalRowData(null, null).build();
        assertEquals("", result.get(Validator.ValidationResource.ROW_DATA));
        assertEquals("", result.get(Validator.ValidationResource.CSV_FILE));
    }

    @Test
    void testBuilderConditionalRowDataRowsPresent() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalRowData(Arrays.asList("row1"), null).build();
        assertFalse(result.containsKey(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testBuilderConditionalRowDataCsvPresent() {
        File mockFile = mock(File.class);
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalRowData(null, Collections.singletonList(mockFile)).build();
        assertFalse(result.containsKey(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testBuilderConditionalRowDataAllNullRows() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalRowData(Arrays.asList((String) null), null).build();
        assertEquals("", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testBuilderConditionalKeyRowsBothMissing() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalKeyRows(null, null).build();
        assertEquals("", result.get(Validator.ValidationResource.KEY_ROWS));
        assertEquals("", result.get(Validator.ValidationResource.CSV_FILE));
    }

    @Test
    void testBuilderConditionalKeyRowsPresent() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalKeyRows(Arrays.asList("key1"), null).build();
        assertFalse(result.containsKey(Validator.ValidationResource.KEY_ROWS));
    }

    @Test
    void testBuilderConditionalColumnsMissing() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalColumns(null).build();
        assertEquals("", result.get(Validator.ValidationResource.COLUMNS));
    }

    @Test
    void testBuilderConditionalColumnsPresent() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalColumns(Arrays.asList("col1")).build();
        assertFalse(result.containsKey(Validator.ValidationResource.COLUMNS));
    }

    @Test
    void testBuilderConditionalColumnsAllNull() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalColumns(Arrays.asList((String) null)).build();
        assertEquals("", result.get(Validator.ValidationResource.COLUMNS));
    }

    @Test
    void testBuilderConditionalLookupRowsMissing() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalLookupRows(null).build();
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_ROWS));
    }

    @Test
    void testBuilderConditionalLookupRowsEmpty() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalLookupRows(Collections.emptyList()).build();
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_ROWS));
    }

    @Test
    void testBuilderConditionalLookupRowsPresent() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .conditionalLookupRows(Arrays.asList("row1")).build();
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_ROWS));
    }

    @Test
    void testBuilderCustom() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.builder()
                .custom(b -> b.fileId("customFile"))
                .build();
        assertEquals("customFile", result.get(Validator.ValidationResource.FILE_ID));
    }

    // Factory method tests
    @Test
    void testCreateAddRowsToTableValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowsToTableValidationMap(
                "drive1", "file1", "table1", null, null);
        assertEquals("drive1", result.get(Validator.ValidationResource.DRIVE_ID));
        assertEquals("file1", result.get(Validator.ValidationResource.FILE_ID));
        assertEquals("table1", result.get(Validator.ValidationResource.TABLE_NAME));
        assertEquals("", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testCreateAddRowsToTableValidationMapForTable() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowsToTableValidationMapForTable(
                "file1", "table1", null, null, 1.0);
        assertEquals("file1", result.get(Validator.ValidationResource.FILE_ID));
        assertEquals("1.0", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testCreateAddRowsToSheetValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowsToSheetValidationMap(
                "drive1", "file1", "Sheet1", Arrays.asList("r1"), null);
        assertEquals("Sheet1", result.get(Validator.ValidationResource.SHEET_NAME));
        assertFalse(result.containsKey(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testCreateRangeValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createRangeValidationMap(
                "drive1", "file1", "Sheet1", "A1:B5");
        assertEquals("A1:B5", result.get(Validator.ValidationResource.RANGE));
    }

    @Test
    void testCreateVlookupValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createVlookupValidationMap(
                "drive1", "file1", "Sheet1", "A1:C5", 2.0);
        assertEquals("2.0", result.get(Validator.ValidationResource.COLUMN_INDEX));
    }

    @Test
    void testCreateColumnsSumValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createColumnsSumValidationMap(
                "drive1", "file1", "table1", null);
        assertEquals("", result.get(Validator.ValidationResource.COLUMNS));
    }

    @Test
    void testCreateVlookupInTableValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createVlookupInTableValidationMap(
                "drive1", "file1", "table1", "val", 1.0);
        assertEquals("val", result.get(Validator.ValidationResource.LOOKUP_VALUE));
    }

    @Test
    void testCreateTableQueryValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createTableQueryValidationMap(
                "drive1", "file1", "table1");
        assertEquals("table1", result.get(Validator.ValidationResource.TABLE_NAME));
    }

    @Test
    void testCreateTableQueryValidationMapForTable() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createTableQueryValidationMapForTable(
                "file1", "table1");
        assertFalse(result.containsKey(Validator.ValidationResource.DRIVE_ID));
    }

    @Test
    void testCreateCellValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createCellValidationMap(
                "drive1", "file1", "Sheet1", "A1");
        assertEquals("A1", result.get(Validator.ValidationResource.CELL_ID));
    }

    @Test
    void testCreateDeleteColumnValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createDeleteColumnValidationMap(
                "drive1", "file1", "table1", "col1");
        assertEquals("col1", result.get(Validator.ValidationResource.COLUMN_NAME));
    }

    @Test
    void testCreateDeleteColumnWithIndexValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createDeleteColumnWithIndexValidationMap(
                "drive1", "file1", "Sheet1", "col1");
        assertEquals("col1", result.get(Validator.ValidationResource.COLUMN_NAME));
    }

    @Test
    void testCreateColumnValuesForKeysValidationMapWithMissingLookup() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createColumnValuesForKeysValidationMap(
                "drive1", "file1", "table1", null, null);
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_ROWS));
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateColumnValuesForKeysValidationMapWithPresentLookup() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createColumnValuesForKeysValidationMap(
                "drive1", "file1", "table1", Arrays.asList("k1"), Arrays.asList("c1"));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_ROWS));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateColumnValuesForKeysValidationMapWithAllNullLookup() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createColumnValuesForKeysValidationMap(
                "drive1", "file1", "table1", Arrays.asList("k1"), Arrays.asList((String) null));
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateColumnValuesForKeysValidationMapForTable() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createColumnValuesForKeysValidationMapForTable(
                "file1", "table1", null, null);
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_ROWS));
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateColumnValuesForKeysValidationMapForTableWithValues() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createColumnValuesForKeysValidationMapForTable(
                "file1", "table1", Arrays.asList("k1"), Arrays.asList("c1"));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_ROWS));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateExecuteFormulaValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createExecuteFormulaValidationMap(
                "drive1", "file1", "=SUM(A1)", "SUM", "total");
        assertEquals("=SUM(A1)", result.get(Validator.ValidationResource.FORMULA));
        assertEquals("SUM", result.get(Validator.ValidationResource.FORMULA_NAME));
        assertEquals("total", result.get(Validator.ValidationResource.COMMENT));
    }

    @Test
    void testCreateSearchValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createSearchValidationMap(
                "drive1", "file1", "Sheet1");
        assertEquals("Sheet1", result.get(Validator.ValidationResource.SHEET_NAME));
    }

    @Test
    void testCreateCellDataValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createCellDataValidationMap(
                "drive1", "file1", "Sheet1", "A1");
        assertEquals("A1", result.get(Validator.ValidationResource.CELL_ID));
    }

    @Test
    void testCreateGetSharedExcelFileValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetSharedExcelFileValidationMap("file.xlsx");
        assertEquals("file.xlsx", result.get(Validator.ValidationResource.SHARED_FILE_NAME));
    }

    @Test
    void testCreateConvertToPdfValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createConvertToPdfValidationMap("drive1", "file1");
        assertEquals("drive1", result.get(Validator.ValidationResource.DRIVE_ID));
        assertEquals("file1", result.get(Validator.ValidationResource.FILE_ID));
    }

    @Test
    void testCreateConvertToPdfValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createConvertToPdfValidationMapForSheet("file1");
        assertEquals("file1", result.get(Validator.ValidationResource.FILE_ID));
        assertFalse(result.containsKey(Validator.ValidationResource.DRIVE_ID));
    }

    @Test
    void testCreateAddRowWithIndexValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowWithIndexValidationMap(
                "drive1", "file1", "Sheet1", null, null, 1.0);
        assertEquals("1.0", result.get(Validator.ValidationResource.ROW_INDEX));
        assertEquals("", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testCreateUpdateRowWithIndexValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createUpdateRowWithIndexValidationMap(
                "drive1", "file1", "Sheet1", "rowData", 2.0);
        assertEquals("rowData", result.get(Validator.ValidationResource.ROW_DATA));
        assertEquals("2.0", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testCreateUpdateRowWithIndexValidationMapNullRow() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createUpdateRowWithIndexValidationMap(
                "drive1", "file1", "Sheet1", null, 2.0);
        assertEquals("", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testCreateGetCompleteRowInfoValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetCompleteRowInfoValidationMap(
                "drive1", "file1", "Sheet1", 3.0);
        assertEquals("3.0", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testCreateGetCompleteRowInfoValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetCompleteRowInfoValidationMapForSheet(
                "file1", "Sheet1", 3.0);
        assertFalse(result.containsKey(Validator.ValidationResource.DRIVE_ID));
    }

    @Test
    void testCreateDeleteRowWithIndexValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createDeleteRowWithIndexValidationMap(
                "drive1", "file1", "Sheet1", 5.0);
        assertEquals("5.0", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testCreateAddMultipleCellDataValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddMultipleCellDataValidationMap(
                "drive1", "file1", "Sheet1");
        assertEquals("Sheet1", result.get(Validator.ValidationResource.SHEET_NAME));
    }

    @Test
    void testCreateAddMultipleCellDataValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddMultipleCellDataValidationMapForSheet(
                "file1", "Sheet1");
        assertFalse(result.containsKey(Validator.ValidationResource.DRIVE_ID));
    }

    @Test
    void testCreateGetColumnValuesForKeysInSheetValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetColumnValuesForKeysInSheetValidationMap(
                "drive1", "file1", "Sheet1", "A1:C5", null, null);
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_ROWS));
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateGetColumnValuesForKeysInSheetValidationMapWithValues() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetColumnValuesForKeysInSheetValidationMap(
                "drive1", "file1", "Sheet1", "A1:C5", Arrays.asList("r1"), Arrays.asList("c1"));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_ROWS));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateGetColumnValuesForKeysInSheetValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetColumnValuesForKeysInSheetValidationMapForSheet(
                "file1", "Sheet1", "A1:C5", null, null);
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_ROWS));
        assertEquals("", result.get(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateGetColumnValuesForKeysInSheetValidationMapForSheetWithValues() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetColumnValuesForKeysInSheetValidationMapForSheet(
                "file1", "Sheet1", "A1:C5", Arrays.asList("k1"), Arrays.asList("c1"));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_ROWS));
        assertFalse(result.containsKey(Validator.ValidationResource.LOOKUP_COLUMNS));
    }

    @Test
    void testCreateClearSharedSheetForRangeValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createClearSharedSheetForRangeValidationMap(
                "drive1", "file1", "Sheet1", "A1:Z100");
        assertEquals("A1:Z100", result.get(Validator.ValidationResource.RANGE));
    }

    @Test
    void testCreateDeleteSheetValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createDeleteSheetValidationMap("file1", "Sheet1");
        assertEquals("Sheet1", result.get(Validator.ValidationResource.SHEET_NAME));
    }

    @Test
    void testCreateGetCellValueValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetCellValueValidationMap("file1", "Sheet1", "A1");
        assertEquals("A1", result.get(Validator.ValidationResource.CELL_ID));
    }

    @Test
    void testCreateAddRowsValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowsValidationMap(
                "file1", "Sheet1", null, null);
        assertEquals("", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testCreateAddRowWithIndexValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowWithIndexValidationMapForSheet(
                "file1", "Sheet1", null, null, 1.0);
        assertEquals("1.0", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testCreateUpdateRowWithIndexValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createUpdateRowWithIndexValidationMapForSheet(
                "file1", "Sheet1", "rowData", 2.0);
        assertEquals("rowData", result.get(Validator.ValidationResource.ROW_DATA));
    }

    @Test
    void testCreateDeleteColumnWithIndexValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createDeleteColumnWithIndexValidationMapForSheet(
                "file1", "Sheet1", "col1");
        assertEquals("col1", result.get(Validator.ValidationResource.COLUMN_NAME));
    }

    @Test
    void testCreateDeleteRowWithIndexValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createDeleteRowWithIndexValidationMapForSheet(
                "file1", "Sheet1", 5.0);
        assertEquals("5.0", result.get(Validator.ValidationResource.ROW_INDEX));
    }

    @Test
    void testCreateFindValueValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createFindValueValidationMapForSheet(
                "file1", "Sheet1", "val");
        assertEquals("val", result.get(Validator.ValidationResource.LOOKUP_VALUE));
    }

    @Test
    void testCreateGetExcelFileInfoValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetExcelFileInfoValidationMap("/path", "file.xlsx");
        assertEquals("/path", result.get(Validator.ValidationResource.PATH));
        assertEquals("file.xlsx", result.get(Validator.ValidationResource.NAME));
    }

    @Test
    void testCreateGetExcelFileInfoValidationMapNulls() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetExcelFileInfoValidationMap(null, null);
        assertEquals("", result.get(Validator.ValidationResource.PATH));
        assertEquals("", result.get(Validator.ValidationResource.NAME));
    }

    @Test
    void testCreateGetSheetsInfoValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createGetSheetsInfoValidationMap("file1");
        assertEquals("file1", result.get(Validator.ValidationResource.FILE_ID));
    }

    @Test
    void testCreateAddDataToCellValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddDataToCellValidationMap(
                "file1", "Sheet1", "A1");
        assertEquals("A1", result.get(Validator.ValidationResource.CELL_ID));
    }

    @Test
    void testCreateExecuteFormulaValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createExecuteFormulaValidationMapForSheet(
                "file1", "=SUM(A1)", "SUM", "total");
        assertEquals("=SUM(A1)", result.get(Validator.ValidationResource.FORMULA));
    }

    @Test
    void testCreateCreateSheetValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createCreateSheetValidationMap("file1");
        assertEquals("file1", result.get(Validator.ValidationResource.FILE_ID));
    }

    @Test
    void testCreateClearSheetForRangeValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createClearSheetForRangeValidationMap(
                "file1", "Sheet1", "A1:Z100");
        assertEquals("A1:Z100", result.get(Validator.ValidationResource.RANGE));
    }

    @Test
    void testCreateVlookupValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createVlookupValidationMapForSheet(
                "file1", "Sheet1", "A1:C5", 2.0, "val");
        assertEquals("val", result.get(Validator.ValidationResource.LOOKUP_VALUE));
    }

    @Test
    void testCreateRangeValidationMapForSheet() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createRangeValidationMapForSheet(
                "file1", "Sheet1", "A1:B5");
        assertEquals("A1:B5", result.get(Validator.ValidationResource.RANGE));
    }

    @Test
    void testCreateAddFormulaToCellValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddFormulaToCellValidationMap(
                "file1", "Sheet1", "A1", "=SUM(B1:B5)");
        assertEquals("=SUM(B1:B5)", result.get(Validator.ValidationResource.FORMULA));
    }

    @Test
    void testCreateMinMaxValueValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createMinMaxValueValidationMap(
                "file1", "Sheet1", "KeyCol", "A:B");
        assertEquals("KeyCol", result.get(Validator.ValidationResource.KEY_COLUMN_NAME));
        assertEquals("A:B", result.get(Validator.ValidationResource.COLUMN_RANGE));
    }

    @Test
    void testCreateMinMaxValueValidationMapNulls() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createMinMaxValueValidationMap(
                "file1", "Sheet1", null, null);
        assertEquals("", result.get(Validator.ValidationResource.KEY_COLUMN_NAME));
        assertEquals("", result.get(Validator.ValidationResource.COLUMN_RANGE));
    }

    @Test
    void testCreateUploadFileAndConvertToPDFValidationMapWithFile() {
        File mockFile = mock(File.class);
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createUploadFileAndConvertToPDFValidationMap(mockFile);
        assertFalse(result.containsKey(Validator.ValidationResource.EXCEL_FILE));
    }

    @Test
    void testCreateUploadFileAndConvertToPDFValidationMapWithNull() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createUploadFileAndConvertToPDFValidationMap(null);
        assertEquals("", result.get(Validator.ValidationResource.EXCEL_FILE));
    }

    @Test
    void testCreateTableVlookupValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createTableVlookupValidationMap(
                "file1", "val", "table1", 1.0);
        assertEquals("val", result.get(Validator.ValidationResource.LOOKUP_VALUE));
        assertEquals("1.0", result.get(Validator.ValidationResource.COLUMN_INDEX));
    }

    @Test
    void testCreateTableCreationValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createTableCreationValidationMap(
                "file1", "Sheet1", "A1:D10");
        assertEquals("A1:D10", result.get(Validator.ValidationResource.RANGE));
    }

    @Test
    void testCreateTableSummationValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createTableSummationValidationMap(
                "file1", "table1", null);
        assertEquals("", result.get(Validator.ValidationResource.COLUMNS));
    }

    @Test
    void testCreateAddRowsInTableValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowsInTableValidationMap(
                "file1", "table1", null, null);
        assertEquals("", result.get(Validator.ValidationResource.KEY_ROWS));
        assertEquals("", result.get(Validator.ValidationResource.CSV_FILE));
    }

    @Test
    void testCreateAddRowsInTableValidationMapWithData() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createAddRowsInTableValidationMap(
                "file1", "table1", Arrays.asList("r1"), null);
        assertFalse(result.containsKey(Validator.ValidationResource.KEY_ROWS));
    }

    @Test
    void testCreateDeleteTableColumnValidationMap() {
        Map<Validator.ValidationResource, String> result = ValidationMapHelper.createDeleteTableColumnValidationMap(
                "file1", "table1", "col1");
        assertEquals("col1", result.get(Validator.ValidationResource.COLUMN_NAME));
    }

    @Test
    void testBuilderBuildReturnsNewMapInstance() {
        ValidationMapHelper.ValidationMapBuilder builder = ValidationMapHelper.builder().fileId("file1");
        Map<Validator.ValidationResource, String> map1 = builder.build();
        Map<Validator.ValidationResource, String> map2 = builder.build();
        assertNotSame(map1, map2);
    }
}
