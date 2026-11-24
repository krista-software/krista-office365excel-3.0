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
import junit.framework.TestCase;
import org.junit.Assert;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class TablesTest extends TestCase {

    @Test()
    public void testGetRowsDataForColumns() {

        FreeForm rowData = new FreeForm();
        List<FreeForm> rowDataList = new ArrayList<>();
        Assert.assertNull(Tables.getRowsDataForColumns(rowDataList, null, List.of(), List.of(), true));
        Assert.assertNull(Tables.getRowsDataForColumns(rowDataList, "", List.of(), List.of(), true));

        // Given Input
        String keyColumnHeader = "name";
        List<String> keyRows = List.of("name1");
        List<String> lookupColumns = List.of("dob","percent","age");

        // File Data
        rowData.put("name","Text", "name1");
        rowData.put("dob", "Text", "12-12-1990");
        rowData.put("age", "Text", "20");
        rowData.put("class", "Text", "5");
        rowData.put("gender", "Text",  "male");
        rowData.put("percent", "Text",  "50");
        rowDataList.add(rowData);

        // Expected Output
        FreeForm expectedOutput = new FreeForm();
        expectedOutput.put("name", "Text", "name1");
        expectedOutput.put("dob", "Text",  "12-12-1990");
        expectedOutput.put("age", "Text",  "20");
        expectedOutput.put("percent", "Text",  "50");

        // Original Output
        List<FreeForm> originalOutput = Tables.getRowsDataForColumns(rowDataList, keyColumnHeader, keyRows, lookupColumns, true);

        Assert.assertEquals(new ArrayList<>(expectedOutput.keySet()), new ArrayList<>(originalOutput.getFirst().keySet()));

    }
}
