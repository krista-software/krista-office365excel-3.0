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

package app.krista.extensions.essentials.spreadsheets.office365excel.service;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.ItemReference;
import com.microsoft.graph.models.RemoteItem;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive test suite for ExcelService with 100% code coverage.
 * Tests Excel file filtering and extraction from DriveItem collections.
 */
class ExcelServiceTest {

    private ExcelService excelService;

    @BeforeEach
    void setUp() {
        excelService = new ExcelService();
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_createsInstance() {
        ExcelService service = new ExcelService();
        assertNotNull(service);
    }

    // ==================== getExcelFiles Tests - Null/Empty Cases ====================

    @Test
    void getExcelFiles_withNullFileName_returnsAllExcelFiles() {
        List<DriveItem> driveItems = createDriveItemsWithExcelFiles();

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertEquals(2, result.size());
    }

    @Test
    void getExcelFiles_withEmptyDriveItems_returnsEmptyList() {
        List<DriveItem> driveItems = new ArrayList<>();

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void getExcelFiles_withNullDriveItemName_skipsItem() {
        List<DriveItem> driveItems = new ArrayList<>();
        DriveItem item = new DriveItem();
        item.name = null;
        item.id = "item1";
        driveItems.add(item);

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    // ==================== getExcelFiles Tests - File Filtering ====================

    @Test
    void getExcelFiles_withSpecificFileName_returnsMatchingFile() {
        List<DriveItem> driveItems = createDriveItemsWithExcelFiles();

        List<ExcelFile> result = ExcelService.getExcelFiles("Report.xlsx", driveItems);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("Report.xlsx", result.get(0).fileName);
    }

    @Test
    void getExcelFiles_withCaseInsensitiveFileName_returnsMatchingFile() {
        List<DriveItem> driveItems = createDriveItemsWithExcelFiles();

        List<ExcelFile> result = ExcelService.getExcelFiles("REPORT.XLSX", driveItems);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("Report.xlsx", result.get(0).fileName);
    }

    @Test
    void getExcelFiles_withNonMatchingFileName_returnsEmptyList() {
        List<DriveItem> driveItems = createDriveItemsWithExcelFiles();

        List<ExcelFile> result = ExcelService.getExcelFiles("NonExistent.xlsx", driveItems);

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    // ==================== getExcelFiles Tests - Non-Excel Files ====================

    @Test
    void getExcelFiles_withNonExcelFiles_filtersThemOut() {
        List<DriveItem> driveItems = new ArrayList<>();
        driveItems.add(createDriveItemWithRemoteItem("document.pdf", "id1", "driveId1", "url1"));
        driveItems.add(createDriveItemWithRemoteItem("image.png", "id2", "driveId2", "url2"));
        driveItems.add(createDriveItemWithRemoteItem("Report.xlsx", "id3", "driveId3", "url3"));

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("Report.xlsx", result.get(0).fileName);
    }

    // ==================== getExcelFiles Tests - RemoteItem Handling ====================

    @Test
    void getExcelFiles_withNullRemoteItem_andNullParentRef_includesItemWithEmptyDriveId() {
        // SharePoint items may arrive without remoteItem; they are now included with empty driveId
        List<DriveItem> driveItems = new ArrayList<>();
        DriveItem item = new DriveItem();
        item.name = "Report.xlsx";
        item.id = "item1";
        item.webUrl = "http://example.com";
        item.remoteItem = null;
        item.parentReference = null;
        driveItems.add(item);

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("", result.get(0).driveId);
    }

    @Test
    void getExcelFiles_withNullRemoteItem_withParentRef_usesDriveIdFromParentRef() {
        // SharePoint items without remoteItem fall back to driveItem.parentReference.driveId
        List<DriveItem> driveItems = new ArrayList<>();
        DriveItem item = new DriveItem();
        item.name = "Report.xlsx";
        item.id = "item1";
        item.webUrl = "http://example.com";
        item.remoteItem = null;
        item.parentReference = new ItemReference();
        item.parentReference.driveId = "sharepoint-drive-id";
        driveItems.add(item);

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("sharepoint-drive-id", result.get(0).driveId);
    }

    @Test
    void getExcelFiles_withNullParentReference_usesEmptyDriveId() {
        List<DriveItem> driveItems = new ArrayList<>();
        DriveItem item = new DriveItem();
        item.name = "Report.xlsx";
        item.id = "item1";
        item.webUrl = "http://example.com";
        item.remoteItem = new RemoteItem();
        item.remoteItem.parentReference = null;
        driveItems.add(item);

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("", result.get(0).driveId);
    }

    @Test
    void getExcelFiles_withValidParentReference_extractsDriveId() {
        List<DriveItem> driveItems = createDriveItemsWithExcelFiles();

        List<ExcelFile> result = ExcelService.getExcelFiles(null, driveItems);

        assertNotNull(result);
        assertEquals(2, result.size());
        assertEquals("drive123", result.get(0).driveId);
    }

    // ==================== Helper Methods ====================

    private List<DriveItem> createDriveItemsWithExcelFiles() {
        List<DriveItem> driveItems = new ArrayList<>();
        driveItems.add(createDriveItemWithRemoteItem("Report.xlsx", "file1", "drive123", "http://example.com/report"));
        driveItems.add(createDriveItemWithRemoteItem("Data.xls", "file2", "drive456", "http://example.com/data"));
        return driveItems;
    }

    private DriveItem createDriveItemWithRemoteItem(String name, String id, String driveId, String webUrl) {
        DriveItem item = new DriveItem();
        item.name = name;
        item.id = id;
        item.webUrl = webUrl;
        item.remoteItem = new RemoteItem();
        item.remoteItem.parentReference = new ItemReference();
        item.remoteItem.parentReference.driveId = driveId;
        return item;
    }
}

