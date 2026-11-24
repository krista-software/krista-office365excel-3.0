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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelSheet;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.models.WorkbookWorksheet;
import org.junit.Test;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class WorksheetsTest {

    public static final String SHEET_1 = "Sheet1";
    public static final String TEST = "Test";

    @Test
    public void getWorksheetsTest() {
        List<WorkbookWorksheet> worksheets = new ArrayList<>();
        WorkbookWorksheet sheet = new WorkbookWorksheet();
        sheet.id = "1";
        sheet.name = SHEET_1;
        worksheets.add(sheet);
        List<ExcelSheet> excelSheets = Worksheets.getWorksheets(worksheets);
        assertEquals(1, excelSheets.size());
        assertEquals(SHEET_1, excelSheets.getFirst().getSheetName());
    }

    // Test case for getWorksheet
    @Test
    public void getWorksheetTest() {
        WorkbookWorksheet worksheet = new WorkbookWorksheet();
        worksheet.id = "1";
        worksheet.name = TEST;
        ExcelSheet excelSheet = Worksheets.getWorksheet(worksheet);
        assertEquals(TEST, excelSheet.getSheetName());
    }

    // Test case for updateWorkbookRange
    @Test
    public void updateWorkbookRangeTest() {
        // Create a sample WorkbookRange
        WorkbookRange workbookRange = new WorkbookRange();
        workbookRange.address = "Sheet1!A1:B2";

        // Create sample data for update
        List<String> rows = Arrays.asList("1,2", "3,4");

        // Test updateWorkbookRange method
        WorkbookRange updatedRange = Worksheets.updateWorkbookRange(SHEET_1, workbookRange, rows, null);

        assertEquals("Sheet1!A3:B4", updatedRange.address);
        assertNotNull(updatedRange.values);
    }

}
