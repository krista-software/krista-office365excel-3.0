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

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class Office365ExcelResourcesTest {

    @Test
    void testDriveId() {
        assertEquals("Drive Id", Office365ExcelResources.DRIVE_ID);
    }

    @Test
    void testFileId() {
        assertEquals("File Id", Office365ExcelResources.FILE_ID);
    }

    @Test
    void testTableName() {
        assertEquals("Table Name", Office365ExcelResources.TABLE_NAME);
    }

    @Test
    void testStateId() {
        assertEquals("State Id", Office365ExcelResources.STATE_ID);
    }

    @Test
    void testRows() {
        assertEquals("Rows", Office365ExcelResources.ROWS);
    }

    @Test
    void testRowsFromCsvFile() {
        assertEquals("CSV File", Office365ExcelResources.ROWS_FROM_CSV_FILE);
    }

    @Test
    void testRange() {
        assertEquals("Range", Office365ExcelResources.RANGE);
    }

    @Test
    void testLookupValue() {
        assertEquals("Lookup Value", Office365ExcelResources.LOOKUP_VALUE);
    }

    @Test
    void testColumnIndex() {
        assertEquals("Column Index", Office365ExcelResources.COLUMN_INDEX);
    }

    @Test
    void testExactMatch() {
        assertEquals("Exact Match", Office365ExcelResources.EXACT_MATCH);
    }

    @Test
    void testCellId() {
        assertEquals("Cell Id", Office365ExcelResources.CELL_ID);
    }

    @Test
    void testData() {
        assertEquals("Data", Office365ExcelResources.DATA);
    }

    @Test
    void testRowIndex() {
        assertEquals("Row Index", Office365ExcelResources.ROW_INDEX);
    }

    @Test
    void testColumnName() {
        assertEquals("Column Name", Office365ExcelResources.COLUMN_NAME);
    }

    @Test
    void testFormula() {
        assertEquals("Formula", Office365ExcelResources.FORMULA);
    }

    @Test
    void testFormulaName() {
        assertEquals("Formula Name", Office365ExcelResources.FORMULA_NAME);
    }

    @Test
    void testComment() {
        assertEquals("Comment", Office365ExcelResources.COMMENT);
    }

    @Test
    void testSheetName() {
        assertEquals("Sheet Name", Office365ExcelResources.SHEET_NAME);
    }

    @Test
    void testDataToSearch() {
        assertEquals("Data to Search", Office365ExcelResources.DATA_TO_SEARCH);
    }

    @Test
    void testSharedFolderName() {
        assertEquals("Shared Folder Name", Office365ExcelResources.SHARED_FOLDER_NAME);
    }

    @Test
    void testSharedFileName() {
        assertEquals("Shared File Name", Office365ExcelResources.SHARED_FILE_NAME);
    }

    @Test
    void testKeyRows() {
        assertEquals("Key Rows", Office365ExcelResources.KEY_ROWS);
    }

    @Test
    void testLookupColumn() {
        assertEquals("Lookup Column", Office365ExcelResources.LOOKUP_COLUMN);
    }

    @Test
    void testGetAllRecord() {
        assertEquals("Get All Record", Office365ExcelResources.GET_ALL_RECORD);
    }

    @Test
    void testColumns() {
        assertEquals("columns", Office365ExcelResources.COLUMNS);
    }

    @Test
    void testRow() {
        assertEquals("Row", Office365ExcelResources.ROW);
    }

    @Test
    void testLookupRows() {
        assertEquals("Lookup Rows", Office365ExcelResources.LOOKUP_ROWS);
    }

    @Test
    void testLookupColumns() {
        assertEquals("Lookup Columns", Office365ExcelResources.LOOKUP_COLUMNS);
    }

    @Test
    void testColumnNumber() {
        assertEquals("Column Number", Office365ExcelResources.COLUMN_NUMBER);
    }

    @Test
    void testCellIdData() {
        assertEquals("Cell ID Data", Office365ExcelResources.CELL_ID_DATA);
    }

    @Test
    void testPath() {
        assertEquals("Path", Office365ExcelResources.PATH);
    }

    @Test
    void testName() {
        assertEquals("Name", Office365ExcelResources.NAME);
    }

    @Test
    void testExcelFile() {
        assertEquals("Excel File", Office365ExcelResources.EXCEL_FILE);
    }

    @Test
    void testFilePath() {
        assertEquals("File Path", Office365ExcelResources.FILE_PATH);
    }

    @Test
    void testColumnRange() {
        assertEquals("Column Range", Office365ExcelResources.COLUMN_RANGE);
    }

    @Test
    void testKeyColumnName() {
        assertEquals("Key Column Name", Office365ExcelResources.KEY_COLUMN_NAME);
    }

    @Test
    void testSeparator() {
        assertEquals("Separator", Office365ExcelResources.SEPARATOR);
    }

    @Test
    void testAuthorization() {
        assertEquals("AUTHORIZATION", Office365ExcelResources.AUTHORIZATION);
    }

    @Test
    void testLogicError() {
        assertEquals("LOGIC_ERROR", Office365ExcelResources.LOGIC_ERROR);
    }

    @Test
    void testSafeStringWithNonNullValue() {
        assertEquals("hello", Office365ExcelResources.safeString("hello"));
    }

    @Test
    void testSafeStringWithNullValue() {
        assertEquals("", Office365ExcelResources.safeString(null));
    }

    @Test
    void testSafeStringWithEmptyString() {
        assertEquals("", Office365ExcelResources.safeString(""));
    }

    @Test
    void testSafeStringWithWhitespace() {
        assertEquals("  ", Office365ExcelResources.safeString("  "));
    }

    @Test
    void testAllConstantsAreNonNull() {
        assertNotNull(Office365ExcelResources.DRIVE_ID);
        assertNotNull(Office365ExcelResources.FILE_ID);
        assertNotNull(Office365ExcelResources.TABLE_NAME);
        assertNotNull(Office365ExcelResources.STATE_ID);
        assertNotNull(Office365ExcelResources.ROWS);
        assertNotNull(Office365ExcelResources.ROWS_FROM_CSV_FILE);
        assertNotNull(Office365ExcelResources.RANGE);
        assertNotNull(Office365ExcelResources.LOOKUP_VALUE);
        assertNotNull(Office365ExcelResources.COLUMN_INDEX);
        assertNotNull(Office365ExcelResources.EXACT_MATCH);
        assertNotNull(Office365ExcelResources.CELL_ID);
        assertNotNull(Office365ExcelResources.DATA);
        assertNotNull(Office365ExcelResources.ROW_INDEX);
        assertNotNull(Office365ExcelResources.COLUMN_NAME);
        assertNotNull(Office365ExcelResources.FORMULA);
        assertNotNull(Office365ExcelResources.FORMULA_NAME);
        assertNotNull(Office365ExcelResources.COMMENT);
        assertNotNull(Office365ExcelResources.SHEET_NAME);
        assertNotNull(Office365ExcelResources.DATA_TO_SEARCH);
        assertNotNull(Office365ExcelResources.SHARED_FOLDER_NAME);
        assertNotNull(Office365ExcelResources.SHARED_FILE_NAME);
        assertNotNull(Office365ExcelResources.KEY_ROWS);
        assertNotNull(Office365ExcelResources.LOOKUP_COLUMN);
        assertNotNull(Office365ExcelResources.GET_ALL_RECORD);
        assertNotNull(Office365ExcelResources.COLUMNS);
        assertNotNull(Office365ExcelResources.ROW);
        assertNotNull(Office365ExcelResources.LOOKUP_ROWS);
        assertNotNull(Office365ExcelResources.LOOKUP_COLUMNS);
        assertNotNull(Office365ExcelResources.COLUMN_NUMBER);
        assertNotNull(Office365ExcelResources.CELL_ID_DATA);
        assertNotNull(Office365ExcelResources.PATH);
        assertNotNull(Office365ExcelResources.NAME);
        assertNotNull(Office365ExcelResources.EXCEL_FILE);
        assertNotNull(Office365ExcelResources.FILE_PATH);
        assertNotNull(Office365ExcelResources.COLUMN_RANGE);
        assertNotNull(Office365ExcelResources.KEY_COLUMN_NAME);
        assertNotNull(Office365ExcelResources.SEPARATOR);
        assertNotNull(Office365ExcelResources.AUTHORIZATION);
        assertNotNull(Office365ExcelResources.LOGIC_ERROR);
    }
}
