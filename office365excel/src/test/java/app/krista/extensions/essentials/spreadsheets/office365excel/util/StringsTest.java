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

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import app.krista.model.base.File;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.junit.Assert.*;

public class StringsTest {
    private StringUtil strings;
    @Mock
    private KristaMediaClient kristaMediaClient;

    @Before
    public void setUp() {
        this.strings = new StringUtil(kristaMediaClient);
    }

    @Test
    public void testIsNullOrEmpty() {
        assertTrue(StringUtil.isNullOrEmpty(""));
        assertTrue(StringUtil.isNullOrEmpty("    "));
        assertFalse(StringUtil.isNullOrEmpty("Hello"));
    }

    @Test
    public void testEmptyIfNull() {
        assertEquals("", StringUtil.emptyIfNull(null));
        assertEquals("", StringUtil.emptyIfNull(""));
        assertEquals("Hello", StringUtil.emptyIfNull("Hello"));
    }

    @Test
    public void testIsExcelFile() {
        assertTrue(StringUtil.isExcelFile("file.xlsx"));
        assertTrue(StringUtil.isExcelFile("file.xls"));
        assertTrue(StringUtil.isExcelFile("file.csv"));
        assertTrue(StringUtil.isExcelFile("file.tsv"));
        assertFalse(StringUtil.isExcelFile("file.txt"));
    }

    @Test
    public void testGetStringList() throws IOException {
        List<String> rows = Arrays.asList("row1", "row2", "row3");
        List<File> files = Arrays.asList(
                new File(10, "file1", "txt", "mediaId1"),
                new File(20, "file2", "txt", "mediaId2"),
                new File(30, "file3", "txt", "mediaId3")
        );
        List<String> result = strings.getStringList(rows, files);
        assertEquals(rows, result);

        result = strings.getStringList(null, Collections.emptyList());
        assertTrue(result.isEmpty());
        List<File> filesFromCsv = Arrays.asList(
                new File(40, "csvFile1", "csv", "csvMediaId1"),
                new File(50, "csvFile2", "csv", "csvMediaId2")
        );
        result = strings.getStringList(rows, filesFromCsv);
        assertEquals("row1", result.get(0));
        assertEquals("row2", result.get(1));
    }

    @Test
    public void testAppendCommas() {
        List<String> inputList = Arrays.asList("a", "1", "x");
        List<String> result = StringUtil.appendCommas(inputList);
        assertEquals("a", result.getFirst());
    }

    @Test
    public void testFindMaxListSize() {
        List<String> inputList = Arrays.asList("a,b,c", "1,2", "x,y,z,w");
        assertEquals(4, StringUtil.findMaxListSize(inputList));
    }

    @Test
    public void testCheckColonPresence() {
        assertTrue(StringUtil.checkColonPresence("abc:def"));
        assertFalse(StringUtil.checkColonPresence("abcdef"));
    }

    @Test
    public void testExtractAlphabetAfterColon() {
        assertEquals("def", StringUtil.extractAlphabetAfterColon("abc:def"));
        assertEquals("", StringUtil.extractAlphabetAfterColon("abcdef"));
    }

    @Test
    public void testRemoveEmptyStringsFromEnd() {
        List<String> inputList = new ArrayList<>();
        inputList.add("a");
        inputList.add(" ");
        StringUtil.removeEmptyStringsFromEnd(inputList);
        assertEquals("a", inputList.getFirst());
    }

    @Test
    public void testCalculateAddressForInsertWithRowsFromCsv() {
        Double rowIndex = 10.0;
        List<String> rows = Arrays.asList("a,b,c", "1,2,", "x,y,z,w");
        String result = strings.calculateAddressForInsert(rowIndex, rows, Collections.emptyList());
        assertEquals("A10:XFD12", result);
    }

    @Test
    public void testCalculateAddressForInsertWithoutRowsFromCsv() {
        Double rowIndex = 10.0;
        List<String> rows = Arrays.asList("a", "1", "x");
        String result = StringUtil.calculateAddressForInsert(rowIndex, rows);
        assertEquals("A10:A12", result);
    }

    @Test
    public void testGetColumnLabel() {
        assertEquals("A", StringUtil.getColumnLabel(1));
        assertEquals("Z", StringUtil.getColumnLabel(26));
        assertEquals("AA", StringUtil.getColumnLabel(27));
        assertEquals("AB", StringUtil.getColumnLabel(28));
        assertEquals("AZ", StringUtil.getColumnLabel(52));
        assertEquals("BA", StringUtil.getColumnLabel(53));
        assertEquals("ZZ", StringUtil.getColumnLabel(702));
        assertEquals("AAA", StringUtil.getColumnLabel(703));
        assertEquals("ZZZ", StringUtil.getColumnLabel(18278));
    }
}
