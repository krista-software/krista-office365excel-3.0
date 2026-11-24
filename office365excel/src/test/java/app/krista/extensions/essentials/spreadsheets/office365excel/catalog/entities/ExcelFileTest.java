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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ExcelFileTest {

    @Test
    void constructor_setsAllFourFields_fieldsAccessible() {
        ExcelFile file = new ExcelFile("file-1", "drive-1", "report.xlsx", "https://example.com/report");

        assertEquals("file-1", file.fileId);
        assertEquals("drive-1", file.driveId);
        assertEquals("report.xlsx", file.fileName);
        assertEquals("https://example.com/report", file.webURL);
    }

    @Test
    void constructor_withNullValues_fieldsAreNull() {
        ExcelFile file = new ExcelFile(null, null, null, null);

        assertNull(file.fileId);
        assertNull(file.driveId);
        assertNull(file.fileName);
        assertNull(file.webURL);
    }

    @Test
    void constructor_withEmptyStrings_fieldsAreEmpty() {
        ExcelFile file = new ExcelFile("", "", "", "");

        assertEquals("", file.fileId);
        assertEquals("", file.driveId);
        assertEquals("", file.fileName);
        assertEquals("", file.webURL);
    }

    @Test
    void fields_arePubliclyAccessible_canBeModified() {
        ExcelFile file = new ExcelFile("f1", "d1", "name1", "url1");

        file.fileId = "f2";
        file.driveId = "d2";
        file.fileName = "name2";
        file.webURL = "url2";

        assertEquals("f2", file.fileId);
        assertEquals("d2", file.driveId);
        assertEquals("name2", file.fileName);
        assertEquals("url2", file.webURL);
    }

    @Test
    void constructor_withMixedNullAndValues_setsCorrectly() {
        ExcelFile file = new ExcelFile("file-2", null, "data.xlsx", null);

        assertEquals("file-2", file.fileId);
        assertNull(file.driveId);
        assertEquals("data.xlsx", file.fileName);
        assertNull(file.webURL);
    }
}
