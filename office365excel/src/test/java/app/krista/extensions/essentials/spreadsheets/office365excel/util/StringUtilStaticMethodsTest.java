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

package app.krista.extensions.essentials.spreadsheets.office365excel.util;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.NullAndEmptySource;
import org.junit.jupiter.params.provider.ValueSource;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive unit tests for all public static methods in StringUtil class
 */
@DisplayName("StringUtil Static Methods Test Suite")
class StringUtilStaticMethodsTest {

    // ==================== isNullOrEmpty() Tests ====================

    @ParameterizedTest
    @NullAndEmptySource
    @ValueSource(strings = {"  ", "\t", "\n", "   \t\n   "})
    @DisplayName("isNullOrEmpty - Null, empty, or whitespace strings")
    void testIsNullOrEmpty_NullEmptyOrWhitespace(String input) {
        assertTrue(StringUtil.isNullOrEmpty(input));
    }

    @ParameterizedTest
    @ValueSource(strings = {"Hello", "  Hello  ", "a", "123", "Hello World"})
    @DisplayName("isNullOrEmpty - Non-empty strings")
    void testIsNullOrEmpty_NonEmpty(String input) {
        assertFalse(StringUtil.isNullOrEmpty(input));
    }

    // ==================== emptyIfNull() Tests ====================

    @Test
    @DisplayName("emptyIfNull - Null input returns empty string")
    void testEmptyIfNull_NullInput() {
        assertEquals("", StringUtil.emptyIfNull(null));
    }

    @ParameterizedTest
    @ValueSource(strings = {"Hello", "", "  ", "Test String"})
    @DisplayName("emptyIfNull - Non-null input returns same string")
    void testEmptyIfNull_NonNullInput(String input) {
        assertEquals(input, StringUtil.emptyIfNull(input));
    }

    // ==================== isExcelFile() Tests ====================

    @ParameterizedTest
    @ValueSource(strings = {
        "document.xlsx",
        "report.xls",
        "data.csv",
        "values.tsv"
    })
    @DisplayName("isExcelFile - Valid Excel file extensions (lowercase)")
    void testIsExcelFile_ValidExtensions(String filename) {
        assertTrue(StringUtil.isExcelFile(filename));
    }

    @ParameterizedTest
    @ValueSource(strings = {
        "file.XLSX",
        "FILE.XLS",
        "DATA.CSV",
        "REPORT.TSV"
    })
    @DisplayName("isExcelFile - Uppercase extensions are NOT supported")
    void testIsExcelFile_UppercaseExtensionsNotSupported(String filename) {
        // Note: The implementation only checks lowercase extensions
        assertFalse(StringUtil.isExcelFile(filename),
            "Uppercase extensions are not supported by current implementation");
    }

    @ParameterizedTest
    @ValueSource(strings = {
        "document.pdf",
        "report.docx",
        "image.png",
        "file.txt",
        "noextension",
        "file.xlsm"
    })
    @DisplayName("isExcelFile - Invalid file extensions")
    void testIsExcelFile_InvalidExtensions(String filename) {
        assertFalse(StringUtil.isExcelFile(filename));
    }

    @Test
    @DisplayName("isExcelFile - Full path with Excel extension")
    void testIsExcelFile_FullPath() {
        assertTrue(StringUtil.isExcelFile("/path/to/file.xlsx"));
        assertTrue(StringUtil.isExcelFile("C:\\Documents\\report.xls"));
    }

    // ==================== appendCommas() Tests ====================

    @Test
    @DisplayName("appendCommas - Append commas to match max count")
    void testAppendCommas_AppendToMatch() {
        List<String> input = Arrays.asList(
            "A,B,C,D",
            "E,F",
            "G,H,I"
        );
        
        List<String> result = StringUtil.appendCommas(input);
        
        assertEquals(3, result.size());
        assertEquals("A,B,C,D", result.get(0));
        assertEquals("E,F,,", result.get(1)); // 2 commas appended
        assertEquals("G,H,I,", result.get(2)); // 1 comma appended
    }

    @Test
    @DisplayName("appendCommas - All strings have same comma count")
    void testAppendCommas_SameCommaCount() {
        List<String> input = Arrays.asList(
            "A,B,C",
            "D,E,F",
            "G,H,I"
        );
        
        List<String> result = StringUtil.appendCommas(input);
        
        assertEquals(3, result.size());
        assertEquals("A,B,C", result.get(0));
        assertEquals("D,E,F", result.get(1));
        assertEquals("G,H,I", result.get(2));
    }

    @Test
    @DisplayName("appendCommas - Empty list")
    void testAppendCommas_EmptyList() {
        List<String> result = StringUtil.appendCommas(new ArrayList<>());
        assertTrue(result.isEmpty());
    }

    @Test
    @DisplayName("appendCommas - Single item")
    void testAppendCommas_SingleItem() {
        List<String> input = Arrays.asList("A,B,C");
        List<String> result = StringUtil.appendCommas(input);
        
        assertEquals(1, result.size());
        assertEquals("A,B,C", result.get(0));
    }

    // ==================== findMaxListSize() Tests ====================

    @Test
    @DisplayName("findMaxListSize - Find maximum column count")
    void testFindMaxListSize_VariousSizes() {
        List<String> input = Arrays.asList(
            "A,B,C",           // 3 columns
            "D,E,F,G,H",       // 5 columns
            "I,J"              // 2 columns
        );
        
        int result = StringUtil.findMaxListSize(input);
        assertEquals(5, result);
    }

    @Test
    @DisplayName("findMaxListSize - Handles quoted commas")
    void testFindMaxListSize_QuotedCommas() {
        List<String> input = Arrays.asList(
            "A,B,\"C,D\",E",   // 4 columns (comma in quotes doesn't count)
            "F,G,H"            // 3 columns
        );
        
        int result = StringUtil.findMaxListSize(input);
        assertEquals(4, result);
    }

    @Test
    @DisplayName("findMaxListSize - Empty list")
    void testFindMaxListSize_EmptyList() {
        int result = StringUtil.findMaxListSize(new ArrayList<>());
        assertEquals(0, result);
    }

    @Test
    @DisplayName("findMaxListSize - No commas")
    void testFindMaxListSize_NoCommas() {
        List<String> input = Arrays.asList("A", "B", "C");
        int result = StringUtil.findMaxListSize(input);
        assertEquals(1, result);
    }

    // ==================== checkColonPresence() Tests ====================

    @ParameterizedTest
    @ValueSource(strings = {"A:B", "Sheet1:Sheet2", "A1:B10", ":", ":::"})
    @DisplayName("checkColonPresence - Strings with colon")
    void testCheckColonPresence_WithColon(String input) {
        assertTrue(StringUtil.checkColonPresence(input));
    }

    @ParameterizedTest
    @ValueSource(strings = {"AB", "Sheet1", "A1B10", "", "NoColon"})
    @DisplayName("checkColonPresence - Strings without colon")
    void testCheckColonPresence_WithoutColon(String input) {
        assertFalse(StringUtil.checkColonPresence(input));
    }

    // ==================== extractAlphabetAfterColon() Tests ====================

    @Test
    @DisplayName("extractAlphabetAfterColon - Extract letters after colon")
    void testExtractAlphabetAfterColon_ValidInput() {
        assertEquals("B", StringUtil.extractAlphabetAfterColon("A:B"));
        assertEquals("AB", StringUtil.extractAlphabetAfterColon("Z:AB"));
        assertEquals("XYZ", StringUtil.extractAlphabetAfterColon("A:XYZ123"));
    }

    @Test
    @DisplayName("extractAlphabetAfterColon - No colon")
    void testExtractAlphabetAfterColon_NoColon() {
        assertEquals("", StringUtil.extractAlphabetAfterColon("ABC"));
    }

    @Test
    @DisplayName("extractAlphabetAfterColon - Colon at end")
    void testExtractAlphabetAfterColon_ColonAtEnd() {
        assertEquals("", StringUtil.extractAlphabetAfterColon("ABC:"));
    }

    @Test
    @DisplayName("extractAlphabetAfterColon - Numbers after colon")
    void testExtractAlphabetAfterColon_NumbersAfterColon() {
        assertEquals("", StringUtil.extractAlphabetAfterColon("A:123"));
    }

    @Test
    @DisplayName("extractAlphabetAfterColon - With spaces")
    void testExtractAlphabetAfterColon_WithSpaces() {
        assertEquals("B", StringUtil.extractAlphabetAfterColon("A:  B  "));
    }

    // ==================== removeEmptyStringsFromEnd() Tests ====================

    @Test
    @DisplayName("removeEmptyStringsFromEnd - Remove trailing empty strings")
    void testRemoveEmptyStringsFromEnd_TrailingEmpty() {
        List<String> input = new ArrayList<>(Arrays.asList("A", "B", "C", "", "", ""));
        StringUtil.removeEmptyStringsFromEnd(input);
        
        assertEquals(3, input.size());
        assertEquals("A", input.get(0));
        assertEquals("B", input.get(1));
        assertEquals("C", input.get(2));
    }

    @Test
    @DisplayName("removeEmptyStringsFromEnd - No trailing empty strings")
    void testRemoveEmptyStringsFromEnd_NoTrailingEmpty() {
        List<String> input = new ArrayList<>(Arrays.asList("A", "B", "C"));
        StringUtil.removeEmptyStringsFromEnd(input);
        
        assertEquals(3, input.size());
    }

    @Test
    @DisplayName("removeEmptyStringsFromEnd - Empty strings in middle")
    void testRemoveEmptyStringsFromEnd_EmptyInMiddle() {
        List<String> input = new ArrayList<>(Arrays.asList("A", "", "B", "", ""));
        StringUtil.removeEmptyStringsFromEnd(input);
        
        assertEquals(3, input.size());
        assertEquals("A", input.get(0));
        assertEquals("", input.get(1));
        assertEquals("B", input.get(2));
    }

    @Test
    @DisplayName("removeEmptyStringsFromEnd - All empty strings")
    void testRemoveEmptyStringsFromEnd_AllEmpty() {
        List<String> input = new ArrayList<>(Arrays.asList("", "", ""));
        StringUtil.removeEmptyStringsFromEnd(input);
        
        assertTrue(input.isEmpty());
    }

    @Test
    @DisplayName("removeEmptyStringsFromEnd - Empty list")
    void testRemoveEmptyStringsFromEnd_EmptyList() {
        List<String> input = new ArrayList<>();
        StringUtil.removeEmptyStringsFromEnd(input);
        
        assertTrue(input.isEmpty());
    }

    // ==================== calculateAddressForInsert() Tests ====================

    @Test
    @DisplayName("calculateAddressForInsert - Calculate cell range for insert")
    void testCalculateAddressForInsert_ValidInput() {
        List<String> rows = Arrays.asList("A,B,C", "D,E,F");
        String result = StringUtil.calculateAddressForInsert(1.0, rows);
        
        assertNotNull(result);
        assertTrue(result.contains(":"));
        assertTrue(result.contains("A"));
    }

    @Test
    @DisplayName("calculateAddressForInsert - Single row")
    void testCalculateAddressForInsert_SingleRow() {
        List<String> rows = Arrays.asList("A,B,C");
        String result = StringUtil.calculateAddressForInsert(1.0, rows);
        
        assertNotNull(result);
        assertTrue(result.matches("A1:[A-Z]+1"));
    }

    @Test
    @DisplayName("calculateAddressForInsert - Multiple rows")
    void testCalculateAddressForInsert_MultipleRows() {
        List<String> rows = Arrays.asList("A,B", "C,D", "E,F");
        String result = StringUtil.calculateAddressForInsert(5.0, rows);
        
        assertNotNull(result);
        assertTrue(result.contains("A5"));
        assertTrue(result.contains("7")); // 5 + 3 - 1 = 7
    }

    // ==================== getColumnLabel() Tests ====================

    @Test
    @DisplayName("getColumnLabel - Convert column number to label")
    void testGetColumnLabel_SingleLetter() {
        assertEquals("A", StringUtil.getColumnLabel(1));
        assertEquals("B", StringUtil.getColumnLabel(2));
        assertEquals("Z", StringUtil.getColumnLabel(26));
    }

    @Test
    @DisplayName("getColumnLabel - Two letter labels")
    void testGetColumnLabel_TwoLetters() {
        assertEquals("AA", StringUtil.getColumnLabel(27));
        assertEquals("AB", StringUtil.getColumnLabel(28));
        assertEquals("AZ", StringUtil.getColumnLabel(52));
        assertEquals("BA", StringUtil.getColumnLabel(53));
    }

    @Test
    @DisplayName("getColumnLabel - Three letter labels")
    void testGetColumnLabel_ThreeLetters() {
        assertEquals("AAA", StringUtil.getColumnLabel(703));
        assertEquals("ZZZ", StringUtil.getColumnLabel(18278));
    }

    @Test
    @DisplayName("getColumnLabel - Large numbers")
    void testGetColumnLabel_LargeNumbers() {
        String result = StringUtil.getColumnLabel(1000);
        assertNotNull(result);
        assertFalse(result.isEmpty());
    }

    // ==================== isDataNullOrEmpty() Tests ====================

    @Test
    @DisplayName("isDataNullOrEmpty - Null list")
    void testIsDataNullOrEmpty_NullList() {
        assertTrue(StringUtil.isDataNullOrEmpty(null));
    }

    @Test
    @DisplayName("isDataNullOrEmpty - Empty list")
    void testIsDataNullOrEmpty_EmptyList() {
        assertTrue(StringUtil.isDataNullOrEmpty(new ArrayList<>()));
    }

    // NOTE: Test for non-empty list is commented out because EntityValue requires constructor parameters
    // This should be an integration test

    /*
    @Test
    @DisplayName("isDataNullOrEmpty - Non-empty list")
    void testIsDataNullOrEmpty_NonEmptyList() {
        List<app.krista.model.base.EntityValue> data = new ArrayList<>();
        data.add(new app.krista.model.base.EntityValue("TestEntity", new HashMap<>()));

        assertFalse(StringUtil.isDataNullOrEmpty(data));
    }
    */
}

