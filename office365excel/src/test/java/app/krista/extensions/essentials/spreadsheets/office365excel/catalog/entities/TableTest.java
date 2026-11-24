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

class TableTest {

    @Test
    void threeArgConstructor_setsAllFields_correctly() {
        Table table = new Table("table-1", "SalesTable", "Sheet1");

        assertEquals("table-1", table.id);
        assertEquals("SalesTable", table.name);
        assertEquals("Sheet1", table.sheetName);
    }

    @Test
    void twoArgConstructor_setsIdAndName_sheetNameNull() {
        Table table = new Table("table-2", "DataTable");

        assertEquals("table-2", table.id);
        assertEquals("DataTable", table.name);
        assertNull(table.sheetName);
    }

    @Test
    void defaultConstructor_allFieldsNull() {
        Table table = new Table();

        assertNull(table.id);
        assertNull(table.name);
        assertNull(table.sheetName);
    }

    @Test
    void fields_arePubliclyAccessible_canBeModified() {
        Table table = new Table();

        table.id = "id-set";
        table.name = "name-set";
        table.sheetName = "sheet-set";

        assertEquals("id-set", table.id);
        assertEquals("name-set", table.name);
        assertEquals("sheet-set", table.sheetName);
    }

    @Test
    void threeArgConstructor_withNullValues_fieldsAreNull() {
        Table table = new Table(null, null, null);

        assertNull(table.id);
        assertNull(table.name);
        assertNull(table.sheetName);
    }

    @Test
    void threeArgConstructor_withEmptyStrings_fieldsAreEmpty() {
        Table table = new Table("", "", "");

        assertEquals("", table.id);
        assertEquals("", table.name);
        assertEquals("", table.sheetName);
    }

    @Test
    void twoArgConstructor_delegatesToThreeArgConstructor_sheetNameNull() {
        Table table = new Table("id-x", "name-x");

        assertEquals("id-x", table.id);
        assertEquals("name-x", table.name);
        assertNull(table.sheetName);
    }
}
