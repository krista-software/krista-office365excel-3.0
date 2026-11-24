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

import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonPrimitive;
import com.microsoft.graph.models.WorkbookFunctionsVlookupParameterSet;
import com.microsoft.graph.models.WorkbookWorksheetRangeParameterSet;
import org.junit.Test;
import java.util.ArrayList;
import java.util.List;
import static org.junit.Assert.*;

public class WorkBookTest {

    public static final String EXAMPLE_XLSX = "example.xlsx";
    public static final String LOOKUP_VALUE = "LookupValue";
    public static final String SOME_VALUE = "SomeValue";
    public static final String LOOKUP_VAL = "lookupVal";
    public static final String TYPE_1 = "Type1";
    public static final String KEY_1 = "Key1";
    public static final String NEW_VALUE = "NewValue";
    public static final String VALUE_1 = "Value1";
    public static final String VALUE_2 = "Value2";

    @Test
    public void testGetPath() {
        String result = Workbooks.getPath("workbooks", EXAMPLE_XLSX);
        assertEquals("workbooks/example.xlsx", result);

        result = Workbooks.getPath(null, EXAMPLE_XLSX);
        assertEquals("/example.xlsx", result);
    }

    @Test
    public void testValidateWorkbookName() {
        String result = Workbooks.validateWorkbookName("example");
        assertEquals(EXAMPLE_XLSX, result);

        result = Workbooks.validateWorkbookName(null);
        assertNull(result);
    }


    @Test
    public void testGetFreeForm() {
        FreeForm result = Workbooks.getFreeForm(LOOKUP_VALUE, SOME_VALUE);
        assertNotNull(result);
        assertEquals(SOME_VALUE, result.get(LOOKUP_VALUE));
    }

    @Test
    public void testGetvLookupParameterSet() {
        WorkbookFunctionsVlookupParameterSet result = Workbooks.getvLookupParameterSet(
                new JsonPrimitive(LOOKUP_VAL),
                new JsonPrimitive(1),
                new JsonPrimitive(true),
                new JsonArray()
        );
        assertNotNull(result);
        assertEquals(LOOKUP_VAL, result.lookupValue.getAsString());
        assertEquals(1, result.colIndexNum.getAsInt());
        assert result.rangeLookup != null;
        assertTrue(result.rangeLookup.getAsBoolean());
        assertEquals(1, result.colIndexNum.getAsInt());
    }

    @Test
    public void testGetWorkbookWorksheetRangeParameterSet() {
        WorkbookWorksheetRangeParameterSet result = Workbooks.getWorkbookWorksheetRangeParameterSet("Sheet1!A1:B2");
        assertNotNull(result);
        assertEquals("Sheet1!A1:B2", result.address);
    }

    @Test
    public void testGetFreeForms() {
        FreeForm finalData = new FreeForm();
        finalData.put(KEY_1, TYPE_1, VALUE_1);

        FreeForm newData = new FreeForm();
        newData.put(KEY_1, "Type2", NEW_VALUE);

        List<FreeForm> finalDataList = new ArrayList<>();
        finalDataList.add(finalData);

        List<FreeForm> newDataList = new ArrayList<>();
        newDataList.add(newData);

        List<FreeForm> result = Workbooks.getFreeForms(finalDataList, newDataList);

        assertEquals(1, result.size());
        assertEquals(TYPE_1, result.getFirst().getTypeDescription(KEY_1));
        assertEquals(NEW_VALUE, result.getFirst().get(KEY_1));
    }

    @Test
    public void testExtractRequiredData() {
        FreeForm finalRow = new FreeForm();
        List<String> dummyRows = new ArrayList<>();
        JsonArray jsonArray = new JsonArray();
        jsonArray.add(new JsonPrimitive(VALUE_1));
        jsonArray.add(new JsonPrimitive(VALUE_2));

        Workbooks.extractRequiredData(finalRow, dummyRows, jsonArray);

        assertEquals(2, dummyRows.size());
        assertEquals(VALUE_1, dummyRows.get(0));
        assertEquals(VALUE_2, dummyRows.get(1));

        assertEquals(VALUE_1, finalRow.get("A"));
        assertEquals(VALUE_2, finalRow.get("B"));
    }

    @Test
    public void testExtractRange() {
        String result = Workbooks.extractRange("Sheet1!A1:B2");
        assertEquals("A1:B2", result);
    }
}
