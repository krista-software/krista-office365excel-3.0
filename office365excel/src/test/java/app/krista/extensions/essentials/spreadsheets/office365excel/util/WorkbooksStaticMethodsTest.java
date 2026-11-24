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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonPrimitive;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.ItemReference;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullAndEmptySource;
import org.junit.jupiter.params.provider.ValueSource;

import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive unit tests for all public static methods in Workbooks class
 */
@DisplayName("Workbooks Static Methods Test Suite")
class WorkbooksStaticMethodsTest {

    // ==================== getPath() Tests ====================
    
    @Test
    @DisplayName("getPath - Valid folder path and workbook name")
    void testGetPath_ValidInputs() {
        String result = Workbooks.getPath("/Documents/Reports", "SalesData");
        assertEquals("/Documents/Reports/SalesData.xlsx", result);
    }

    @Test
    @DisplayName("getPath - Null folder path")
    void testGetPath_NullFolderPath() {
        String result = Workbooks.getPath(null, "SalesData");
        assertEquals("/SalesData.xlsx", result);
    }

    @Test
    @DisplayName("getPath - Empty folder path")
    void testGetPath_EmptyFolderPath() {
        String result = Workbooks.getPath("", "SalesData");
        assertEquals("/SalesData.xlsx", result);
    }

    @Test
    @DisplayName("getPath - Workbook name already has extension")
    void testGetPath_WorkbookNameWithExtension() {
        String result = Workbooks.getPath("/Documents", "SalesData.xlsx");
        assertEquals("/Documents/SalesData.xlsx", result);
    }

    // ==================== validateWorkbookName() Tests ====================

    @Test
    @DisplayName("validateWorkbookName - Valid name without extension")
    void testValidateWorkbookName_WithoutExtension() {
        String result = Workbooks.validateWorkbookName("SalesData");
        assertEquals("SalesData.xlsx", result);
    }

    @Test
    @DisplayName("validateWorkbookName - Valid name with extension")
    void testValidateWorkbookName_WithExtension() {
        String result = Workbooks.validateWorkbookName("SalesData.xlsx");
        assertEquals("SalesData.xlsx", result);
    }

    @ParameterizedTest
    @NullAndEmptySource
    @ValueSource(strings = {"  ", "\t", "\n"})
    @DisplayName("validateWorkbookName - Null or empty inputs")
    void testValidateWorkbookName_NullOrEmpty(String input) {
        String result = Workbooks.validateWorkbookName(input);
        assertNull(result);
    }

    @Test
    @DisplayName("validateWorkbookName - Name with special characters")
    void testValidateWorkbookName_SpecialCharacters() {
        String result = Workbooks.validateWorkbookName("Sales_Data-2024");
        assertEquals("Sales_Data-2024.xlsx", result);
    }

    // ==================== getWorkbooks() Tests ====================

    @Test
    @DisplayName("getWorkbooks - Find matching workbooks")
    void testGetWorkbooks_MatchingWorkbooks() {
        List<DriveItem> driveItems = createTestDriveItems();
        List<ExcelFile> result = Workbooks.getWorkbooks("Sales", driveItems);
        
        assertEquals(2, result.size());
        assertTrue(result.stream().anyMatch(f -> f.fileName.equals("SalesData.xlsx")));
        assertTrue(result.stream().anyMatch(f -> f.fileName.equals("Sales_Report.xlsx")));
    }

    @Test
    @DisplayName("getWorkbooks - Case insensitive search")
    void testGetWorkbooks_CaseInsensitive() {
        List<DriveItem> driveItems = createTestDriveItems();
        List<ExcelFile> result = Workbooks.getWorkbooks("sales", driveItems);
        
        assertEquals(2, result.size());
    }

    @Test
    @DisplayName("getWorkbooks - No matching workbooks")
    void testGetWorkbooks_NoMatches() {
        List<DriveItem> driveItems = createTestDriveItems();
        List<ExcelFile> result = Workbooks.getWorkbooks("NonExistent", driveItems);
        
        assertTrue(result.isEmpty());
    }

    @Test
    @DisplayName("getWorkbooks - Empty drive items list")
    void testGetWorkbooks_EmptyList() {
        List<ExcelFile> result = Workbooks.getWorkbooks("Sales", new ArrayList<>());
        assertTrue(result.isEmpty());
    }

    @Test
    @DisplayName("getWorkbooks - Filters out non-Excel files")
    void testGetWorkbooks_FiltersNonExcelFiles() {
        List<DriveItem> driveItems = createMixedFileTypesDriveItems();
        List<ExcelFile> result = Workbooks.getWorkbooks("Report", driveItems);
        
        // Should only include .xlsx files, not .pdf or .docx
        assertEquals(1, result.size());
        assertTrue(result.get(0).fileName.endsWith(".xlsx"));
    }

    // ==================== getFreeForm() Tests ====================

    @Test
    @DisplayName("getFreeForm - Valid lookup value and value")
    void testGetFreeForm_ValidInputs() {
        FreeForm result = Workbooks.getFreeForm("ProductName", "Laptop");
        
        assertNotNull(result);
        assertEquals("Laptop", result.get("ProductName"));
        assertEquals(Constants.TEXT, result.getTypeDescription("ProductName"));
    }

    @Test
    @DisplayName("getFreeForm - Null value")
    void testGetFreeForm_NullValue() {
        FreeForm result = Workbooks.getFreeForm("ProductName", null);
        
        assertNotNull(result);
        assertEquals("", result.get("ProductName"));
    }

    @Test
    @DisplayName("getFreeForm - Numeric value")
    void testGetFreeForm_NumericValue() {
        FreeForm result = Workbooks.getFreeForm("Price", 99.99);
        
        assertNotNull(result);
        assertEquals(99.99, result.get("Price"));
    }

    // ==================== extractRange() Tests ====================

    @Test
    @DisplayName("extractRange - Valid range format")
    void testExtractRange_ValidFormat() {
        String result = Workbooks.extractRange("Sheet1!A1:C10");
        assertEquals("A1:C10", result);
    }

    @Test
    @DisplayName("extractRange - Invalid format without exclamation")
    void testExtractRange_InvalidFormat_NoExclamation() {
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            Workbooks.extractRange("A1:C10");
        });
        assertTrue(exception.getMessage().contains("Invalid input format"));
    }

    @Test
    @DisplayName("extractRange - Invalid format with multiple exclamations")
    void testExtractRange_InvalidFormat_MultipleExclamations() {
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            Workbooks.extractRange("Sheet1!Sheet2!A1:C10");
        });
        assertTrue(exception.getMessage().contains("Invalid input format"));
    }

    @ParameterizedTest
    @CsvSource({
        "'Sales!A1:Z100', 'A1:Z100'",
        "'Data Sheet!B2:D50', 'B2:D50'",
        "'Sheet_1!AA1:ZZ999', 'AA1:ZZ999'"
    })
    @DisplayName("extractRange - Various valid formats")
    void testExtractRange_VariousValidFormats(String input, String expected) {
        String result = Workbooks.extractRange(input);
        assertEquals(expected, result);
    }

    // ==================== getFreeForms() Tests ====================

    @Test
    @DisplayName("getFreeForms - Merge new data with existing data")
    void testGetFreeForms_MergeData() {
        List<FreeForm> finalData = createFinalDataList();
        List<FreeForm> newData = createNewDataList();
        
        List<FreeForm> result = Workbooks.getFreeForms(finalData, newData);
        
        assertNotNull(result);
        assertFalse(result.isEmpty());
        // Verify data types are preserved from finalData
        for (FreeForm form : result) {
            for (String key : form.keySet()) {
                if (finalData.getFirst().containsKey(key)) {
                    assertEquals(finalData.getFirst().getTypeDescription(key), 
                               form.getTypeDescription(key));
                }
            }
        }
    }

    @Test
    @DisplayName("getFreeForms - Empty final data")
    void testGetFreeForms_EmptyFinalData() {
        List<FreeForm> finalData = new ArrayList<>();
        List<FreeForm> newData = createNewDataList();
        
        List<FreeForm> result = Workbooks.getFreeForms(finalData, newData);
        
        assertNotNull(result);
        assertEquals(newData.size(), result.size());
    }

    @Test
    @DisplayName("getFreeForms - Empty new data")
    void testGetFreeForms_EmptyNewData() {
        List<FreeForm> finalData = createFinalDataList();
        List<FreeForm> newData = new ArrayList<>();
        
        List<FreeForm> result = Workbooks.getFreeForms(finalData, newData);
        
        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    // ==================== extractRequiredData() Tests ====================

    @Test
    @DisplayName("extractRequiredData - Extract data from JsonArray")
    void testExtractRequiredData_ValidData() {
        FreeForm finalRow = new FreeForm();
        List<String> dummyRows = new ArrayList<>();
        JsonArray jsonArray = createTestJsonArray("Value1", "Value2", "Value3");
        
        Workbooks.extractRequiredData(finalRow, dummyRows, jsonArray);
        
        assertEquals(3, dummyRows.size());
        assertEquals("Value1", dummyRows.get(0));
        assertEquals("Value2", dummyRows.get(1));
        assertEquals("Value3", dummyRows.get(2));
    }

    @Test
    @DisplayName("extractRequiredData - Empty JsonArray")
    void testExtractRequiredData_EmptyArray() {
        FreeForm finalRow = new FreeForm();
        List<String> dummyRows = new ArrayList<>();
        JsonArray jsonArray = new JsonArray();
        
        Workbooks.extractRequiredData(finalRow, dummyRows, jsonArray);
        
        assertTrue(dummyRows.isEmpty());
    }

    @Test
    @DisplayName("extractRequiredData - Removes trailing empty strings")
    void testExtractRequiredData_RemovesTrailingEmpty() {
        FreeForm finalRow = new FreeForm();
        List<String> dummyRows = new ArrayList<>();
        JsonArray jsonArray = createTestJsonArray("Value1", "Value2", "", "");
        
        Workbooks.extractRequiredData(finalRow, dummyRows, jsonArray);
        
        // Should remove trailing empty strings
        assertEquals(2, dummyRows.size());
    }

    // ==================== Helper Methods ====================

    private List<DriveItem> createTestDriveItems() {
        List<DriveItem> items = new ArrayList<>();
        
        items.add(createDriveItem("SalesData.xlsx", "id1", "driveId1"));
        items.add(createDriveItem("Sales_Report.xlsx", "id2", "driveId2"));
        items.add(createDriveItem("Inventory.xlsx", "id3", "driveId3"));
        
        return items;
    }

    private List<DriveItem> createMixedFileTypesDriveItems() {
        List<DriveItem> items = new ArrayList<>();
        
        items.add(createDriveItem("Report.xlsx", "id1", "driveId1"));
        items.add(createDriveItem("Report.pdf", "id2", "driveId2"));
        items.add(createDriveItem("Report.docx", "id3", "driveId3"));
        
        return items;
    }

    private DriveItem createDriveItem(String name, String id, String driveId) {
        DriveItem item = new DriveItem();
        item.name = name;
        item.id = id;
        item.webUrl = "https://example.com/" + name;
        
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = driveId;
        item.parentReference = parentRef;
        
        return item;
    }

    private List<FreeForm> createFinalDataList() {
        List<FreeForm> list = new ArrayList<>();
        FreeForm form = new FreeForm();
        form.put("Name", Constants.TEXT, "Product1");
        form.put("Price", "Number", 100.0);
        list.add(form);
        return list;
    }

    private List<FreeForm> createNewDataList() {
        List<FreeForm> list = new ArrayList<>();
        FreeForm form = new FreeForm();
        form.put("Name", Constants.TEXT, "Product2");
        form.put("Price", Constants.TEXT, "200");
        list.add(form);
        return list;
    }

    private JsonArray createTestJsonArray(String... values) {
        JsonArray array = new JsonArray();
        for (String value : values) {
            array.add(new JsonPrimitive(value));
        }
        return array;
    }
}

