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

class ExcelSheetTest {

    @Test
    void constructor_setsFields_correctly() {
        ExcelSheet sheet = new ExcelSheet("sheet-1", "Sales Data");

        assertEquals("Sales Data", sheet.getSheetName());
    }

    @Test
    void constructor_withNullValues_fieldsAreNull() {
        ExcelSheet sheet = new ExcelSheet(null, null);

        assertNull(sheet.getSheetName());
    }

    @Test
    void constructor_withEmptyStrings_fieldsAreEmpty() {
        ExcelSheet sheet = new ExcelSheet("", "");

        assertEquals("", sheet.getSheetName());
    }

    @Test
    void getSheetName_returnsCorrectValue() {
        ExcelSheet sheet = new ExcelSheet("id-abc", "Sheet1");

        assertEquals("Sheet1", sheet.getSheetName());
    }

    @Test
    void constructor_setsSheetId_fieldAccessibleViaPackage() {
        ExcelSheet sheet = new ExcelSheet("sheet-id-123", "MySheet");

        // sheetId is package-private, accessible within same package
        assertEquals("sheet-id-123", sheet.sheetId);
    }

    @Test
    void constructor_setsSheetName_fieldAccessibleViaPackage() {
        ExcelSheet sheet = new ExcelSheet("id", "WorksheetName");

        // sheetName is package-private, accessible within same package
        assertEquals("WorksheetName", sheet.sheetName);
    }
}
