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

package app.krista.extensions.essentials.spreadsheets.office365excel.util.parser;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;
import com.google.gson.JsonPrimitive;
import com.microsoft.graph.models.WorkbookTable;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class DataParserTest {

    @Test
    void parse_withValidJsonArray_returnsFreeFormList() {
        // [[1, 2], [3, 4]]
        JsonElement element = JsonParser.parseString("[[1, 2], [3, 4]]");

        List<FreeForm> result = DataParser.parse(element);

        assertNotNull(result);
        assertEquals(2, result.size());
    }

    @Test
    void parse_withBlankSheet_returnsEmptyList() {
        // [[""]]: a blank sheet
        JsonElement element = JsonParser.parseString("[[\"\"]]");

        List<FreeForm> result = DataParser.parse(element);

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void parse_withStringValues_returnsFreeFormWithTextType() {
        JsonElement element = JsonParser.parseString("[[\"hello\", \"world\"]]");

        List<FreeForm> result = DataParser.parse(element);

        assertNotNull(result);
        assertEquals(1, result.size());
        FreeForm row = result.get(0);
        assertEquals("hello", row.get("A"));
        assertEquals("world", row.get("B"));
    }

    @Test
    void parse_withNumericValues_returnsFreeFormWithNumberType() {
        JsonElement element = JsonParser.parseString("[[42, 99.5]]");

        List<FreeForm> result = DataParser.parse(element);

        assertNotNull(result);
        assertEquals(1, result.size());
        FreeForm row = result.get(0);
        assertEquals(42.0, row.get("A"));
        assertEquals(99.5, row.get("B"));
    }

    @Test
    void parse_withMixedValues_handlesCorrectly() {
        JsonElement element = JsonParser.parseString("[[\"Name\", 100]]");

        List<FreeForm> result = DataParser.parse(element);

        assertNotNull(result);
        assertEquals(1, result.size());
    }

    @Test
    void getColumnHeaders_extractsHeaders() {
        JsonElement element = JsonParser.parseString("[[\"Name\", \"Age\", \"City\"]]");

        List<String> headers = DataParser.getColumnHeaders(element);

        assertEquals(3, headers.size());
        assertEquals("Name", headers.get(0));
        assertEquals("Age", headers.get(1));
        assertEquals("City", headers.get(2));
    }

    @Test
    void getColumnHeaders_withNullElement_returnsEmptyList() {
        List<String> headers = DataParser.getColumnHeaders(null);

        assertNotNull(headers);
        assertTrue(headers.isEmpty());
    }

    @Test
    void getCellData_withJsonPrimitiveNumber_returnsDouble() {
        JsonPrimitive number = new JsonPrimitive(42);

        Object result = DataParser.getCellData(number);

        assertEquals(42.0, result);
    }

    @Test
    void getCellData_withJsonPrimitiveString_returnsString() {
        JsonPrimitive str = new JsonPrimitive("hello");

        Object result = DataParser.getCellData(str);

        assertEquals("hello", result);
    }

    @Test
    void getCellData_withJsonArray_extractsFirstValue() {
        // [[42]]
        JsonElement element = JsonParser.parseString("[[42]]");

        Object result = DataParser.getCellData(element);

        assertEquals(42.0, result);
    }

    @Test
    void getCellData_withNestedJsonArrayString_extractsStringValue() {
        JsonElement element = JsonParser.parseString("[[\"text\"]]");

        Object result = DataParser.getCellData(element);

        assertEquals("text", result);
    }

    @Test
    void getCellData_withEmptyJsonArray_returnsNull() {
        JsonArray emptyOuter = new JsonArray();

        Object result = DataParser.getCellData(emptyOuter);

        assertNull(result);
    }

    @Test
    void parse_workbookTable_withValidTable_createsTable() {
        WorkbookTable workbookTable = new WorkbookTable();
        workbookTable.id = "table-1";
        workbookTable.name = "SalesTable";

        Table result = DataParser.parse(workbookTable, "Sheet1");

        assertNotNull(result);
        assertEquals("table-1", result.id);
        assertEquals("SalesTable", result.name);
        assertEquals("Sheet1", result.sheetName);
    }

    @Test
    void parse_workbookTable_withNullTable_returnsNull() {
        Table result = DataParser.parse(null, "Sheet1");

        assertNull(result);
    }

    @Test
    void parse_workbookTable_withNullWorksheetName_setsNullSheetName() {
        WorkbookTable workbookTable = new WorkbookTable();
        workbookTable.id = "t-id";
        workbookTable.name = "t-name";

        Table result = DataParser.parse(workbookTable, null);

        assertNotNull(result);
        assertNull(result.sheetName);
    }

    @Test
    void convertRangeIntoFlatRange_convertsNestedArrayToFlat() {
        JsonElement element = JsonParser.parseString("[[1, 2], [3, 4]]");

        JsonElement result = DataParser.convertRangeIntoFlatRange(element);

        assertNotNull(result);
        assertTrue(result.isJsonArray());
        JsonArray array = result.getAsJsonArray();
        assertEquals(4, array.size());
    }

    @Test
    void convertToFreeForm_withNumber_setsNumberType() {
        JsonPrimitive number = new JsonPrimitive(42);

        FreeForm result = DataParser.convertToFreeForm("SUM", number);

        assertNotNull(result);
        assertEquals(42, result.get("SUM"));
        assertEquals("Number", result.getTypeDescription("SUM"));
    }

    @Test
    void convertToFreeForm_withString_setsTextType() {
        JsonPrimitive str = new JsonPrimitive("result text");

        FreeForm result = DataParser.convertToFreeForm("FORMULA", str);

        assertNotNull(result);
        assertEquals("result text", result.get("FORMULA"));
        assertEquals("Text", result.getTypeDescription("FORMULA"));
    }

    @Test
    void convertToFreeForm_withJsonArray_returnsEmptyFreeForm() {
        JsonArray array = new JsonArray();
        array.add(1);

        FreeForm result = DataParser.convertToFreeForm("FORMULA", array);

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void parse_multipleRows_generatesCorrectHeaders() {
        // Tests that getHeader generates A, B, C, ... for columns
        JsonElement element = JsonParser.parseString("[[1, 2, 3, 4, 5]]");

        List<FreeForm> result = DataParser.parse(element);

        assertEquals(1, result.size());
        FreeForm row = result.get(0);
        assertNotNull(row.get("A"));
        assertNotNull(row.get("B"));
        assertNotNull(row.get("C"));
        assertNotNull(row.get("D"));
        assertNotNull(row.get("E"));
    }

    @Test
    void parse_moreThan26Columns_generatesDoubleLetterHeaders() {
        // Build a row with 27 elements to test AA column header generation
        StringBuilder sb = new StringBuilder("[[");
        for (int i = 0; i < 27; i++) {
            if (i > 0) sb.append(", ");
            sb.append(i);
        }
        sb.append("]]");
        JsonElement element = JsonParser.parseString(sb.toString());

        List<FreeForm> result = DataParser.parse(element);

        assertEquals(1, result.size());
        FreeForm row = result.get(0);
        // 27th column (index 26) should be "AA"
        assertNotNull(row.get("AA"));
    }

    @Test
    void parse_multipleRowsWithNumbers_previousRowInfluencesType() {
        // First row numeric, second row numeric -> both should be Number type
        JsonElement element = JsonParser.parseString("[[10], [20]]");

        List<FreeForm> result = DataParser.parse(element);

        assertEquals(2, result.size());
        assertEquals(10.0, result.get(0).get("A"));
        assertEquals(20.0, result.get(1).get("A"));
    }

    @Test
    void getColumnHeaders_withJsonObject_extractsFromTextKey() {
        // When jsonElement is a JsonObject with "text" key
        JsonElement element = JsonParser.parseString("{\"text\": [[\"Header1\", \"Header2\"]]}");

        List<String> headers = DataParser.getColumnHeaders(element);

        assertEquals(2, headers.size());
        assertEquals("Header1", headers.get(0));
        assertEquals("Header2", headers.get(1));
    }
}
