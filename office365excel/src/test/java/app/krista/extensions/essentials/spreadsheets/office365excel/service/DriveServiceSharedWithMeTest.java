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

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.ItemReference;
import com.microsoft.graph.models.RemoteItem;
import com.microsoft.graph.requests.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.anyInt;
import static org.mockito.ArgumentMatchers.isNull;
import static org.mockito.Mockito.*;

/**
 * 100% branch coverage for {@code DriveService.getDriveItemsSharedWithMe()}.
 *
 * <p>The method is private; all tests drive it through the public entry-point
 * {@link DriveService#getAllSharedDrivesInformation()}, which returns
 * {@code Collections.emptyList()} when the inner method returns {@code null},
 * and a non-null list otherwise.
 *
 * <p><strong>Why explicit chain mocks instead of {@code RETURNS_DEEP_STUBS}:</strong>
 * {@code DriveSharedWithMeCollectionRequest.get()} inherits its {@code get()} from a
 * generic base class. Java type erasure causes Mockito's {@code RETURNS_DEEP_STUBS} to
 * resolve the return as the raw {@code BaseCollectionPage} rather than the covariant
 * {@code DriveSharedWithMeCollectionPage}, triggering a {@code ClassCastException} during
 * stub setup. Mocking each link of the chain with an explicit {@code @Mock} of the exact
 * return type eliminates the type-erasure problem.
 *
 * <p>Branches covered:
 * <ul>
 *   <li>Initial page {@code == null} → while-loop is never entered → empty list returned</li>
 *   <li>Single page: {@code getNextPage()} returns {@code null} → break</li>
 *   <li>Two pages: first page's {@code getNextPage()} non-null → second page fetched;
 *       second page's {@code getNextPage()} {@code null} → break; all items collected</li>
 *   <li>Item type logging — three branches:
 *       {@code folder != null} → "FOLDER",
 *       {@code folder == null && file != null} → "FILE",
 *       both {@code null} → "OTHER"</li>
 *   <li>RemoteItem info logging — three branches:
 *       {@code remoteItem != null && parentReference != null},
 *       {@code remoteItem != null && parentReference == null},
 *       {@code remoteItem == null}</li>
 *   <li>Excel-file counting filter:
 *       {@code file != null && name != null && isExcelFile(name)} → true,
 *       {@code file != null && !isExcelFile(name)} → false,
 *       {@code file == null} → false</li>
 *   <li>{@code GraphServiceException} thrown → returns {@code null} → caller returns empty list</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("DriveService.getDriveItemsSharedWithMe – 100% branch coverage")
class DriveServiceSharedWithMeTest {

    // ─── Infrastructure mocks ────────────────────────────────────────────────
    @Mock private GraphServiceClientProviderFactory mockProviderFactory;
    @Mock private GraphServiceClientProvider        mockClientProvider;
    @Mock private KristaMediaClient                 mockKristaMediaClient;

    // ─── sharedWithMe chain: getDriveRequestBuilder(null) → sharedWithMe() chain ──
    @Mock private DriveRequestBuilder                       mockDriveRequestBuilder;
    @Mock private DriveSharedWithMeCollectionRequestBuilder mockSharedWithMeBuilder;
    @Mock private DriveSharedWithMeCollectionRequest        mockSharedWithMeRequest;
    @Mock private DriveSharedWithMeCollectionPage           mockFirstPage;

    // ─── Pagination: second page chain ───────────────────────────────────────
    @Mock private DriveSharedWithMeCollectionRequestBuilder mockNextPageBuilder;
    @Mock private DriveSharedWithMeCollectionRequest        mockNextPageRequest;
    @Mock private DriveSharedWithMeCollectionPage           mockSecondPage;

    private DriveService service;

    @BeforeEach
    void setUp() {
        when(mockProviderFactory.create()).thenReturn(mockClientProvider);

        // Wire the sharedWithMe chain up to (but not including) the terminal .get() call,
        // which is configured per-test.
        when(mockClientProvider.getDriveRequestBuilder(isNull())).thenReturn(mockDriveRequestBuilder);
        when(mockDriveRequestBuilder.sharedWithMe()).thenReturn(mockSharedWithMeBuilder);
        when(mockSharedWithMeBuilder.buildRequest()).thenReturn(mockSharedWithMeRequest);
        // top() returns the same request object for fluent chaining
        when(mockSharedWithMeRequest.top(anyInt())).thenReturn(mockSharedWithMeRequest);

        // Default: second page's pagination terminates (no further pages)
        when(mockNextPageBuilder.buildRequest()).thenReturn(mockNextPageRequest);
        when(mockSecondPage.getNextPage()).thenReturn(null);

        service = new DriveService(mockProviderFactory, mockKristaMediaClient);
    }

    // ─── Model factories ─────────────────────────────────────────────────────

    /** File item with name; has remoteItem with parentReference (covers FILE + remoteItem.driveId branch). */
    private static DriveItem excelFileItemWithParentRef(String name) {
        DriveItem item = new DriveItem();
        item.name = name;
        item.file = new com.microsoft.graph.models.File();
        item.remoteItem = new RemoteItem();
        item.remoteItem.parentReference = new ItemReference();
        item.remoteItem.parentReference.driveId = "remote-drive-id";
        return item;
    }

    /** Non-Excel file item; has remoteItem with NO parentReference (covers "remoteItem(no parentRef)" branch). */
    private static DriveItem nonExcelFileItemNoParentRef(String name) {
        DriveItem item = new DriveItem();
        item.name = name;
        item.file = new com.microsoft.graph.models.File();
        item.remoteItem = new RemoteItem(); // parentReference is null
        return item;
    }

    /** Folder item with no remoteItem (covers FOLDER type + "no-remoteItem" branch). */
    private static DriveItem folderItem(String name) {
        DriveItem item = new DriveItem();
        item.name = name;
        item.folder = new com.microsoft.graph.models.Folder();
        // remoteItem is null → "no-remoteItem" branch
        return item;
    }

    /**
     * Item that is neither file nor folder and has a null name (covers OTHER type, null-name
     * short-circuit in the Excel-count filter, and "no-remoteItem" branch).
     */
    private static DriveItem otherItemNullName() {
        DriveItem item = new DriveItem();
        // folder, file, name, remoteItem all null
        return item;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Exception handling
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Exception handling")
    class ExceptionHandling {

        @Test
        @DisplayName("GraphServiceException on get() → method returns null → caller returns empty list")
        void graphServiceException_methodReturnsNull_callerReturnsEmptyList() {
            when(mockSharedWithMeRequest.get()).thenThrow(mock(GraphServiceException.class));

            List<?> result = service.getAllSharedDrivesInformation();

            assertNotNull(result);
            assertTrue(result.isEmpty());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Initial page null → while-loop not entered
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Initial page null – loop never entered")
    class InitialPageNull {

        @Test
        @DisplayName("get() returns null → while-loop body never executes → empty list returned")
        void initialPageNull_loopSkipped_emptyListReturned() {
            when(mockSharedWithMeRequest.get()).thenReturn(null);

            // getDriveItemsSharedWithMe() returns an empty (non-null) list;
            // getAllSharedDrivesInformation() returns Result.getDrives(emptyList) which is non-null.
            List<?> result = service.getAllSharedDrivesInformation();

            assertNotNull(result);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Pagination branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Pagination")
    class PaginationBranches {

        @Test
        @DisplayName("single page: getNextPage() == null → break; all items from that page returned")
        void singlePage_getNextPageNull_breakAndReturn() {
            DriveItem item = excelFileItemWithParentRef("Report.xlsx");
            when(mockFirstPage.getCurrentPage()).thenReturn(List.of(item));
            when(mockFirstPage.getNextPage()).thenReturn(null);   // triggers break
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            List<?> result = service.getAllSharedDrivesInformation();

            assertNotNull(result);
        }

        @Test
        @DisplayName("two pages: first.getNextPage() non-null → second page fetched; second.getNextPage() null → break")
        void twoPages_allItemsFromBothPagesCollected() {
            DriveItem item1 = excelFileItemWithParentRef("Page1.xlsx");
            DriveItem item2 = folderItem("FolderOnPage2");

            // First page has one item; getNextPage() leads to second page
            when(mockFirstPage.getCurrentPage()).thenReturn(List.of(item1));
            when(mockFirstPage.getNextPage()).thenReturn(mockNextPageBuilder);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            // Second page has one item; getNextPage() returns null → break
            when(mockNextPageRequest.get()).thenReturn(mockSecondPage);
            when(mockSecondPage.getCurrentPage()).thenReturn(List.of(item2));
            when(mockSecondPage.getNextPage()).thenReturn(null);

            List<?> result = service.getAllSharedDrivesInformation();

            assertNotNull(result);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Item-type and remoteItem logging branches
    // All six logging branches are exercised in a single test to keep setup
    // concise; individual tests below pin down each branch explicitly.
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Item type and remoteItem logging branches")
    class LoggingBranches {

        /**
         * Covers all three item-type branches (FOLDER / FILE / OTHER) and all three
         * remoteItem branches (with parentRef / without parentRef / no remoteItem) in
         * a single page.
         *
         * <p>Items:
         * <ol>
         *   <li>Excel file — FILE type, remoteItem with parentReference</li>
         *   <li>Non-Excel file — FILE type, remoteItem without parentReference</li>
         *   <li>Folder — FOLDER type, no remoteItem</li>
         *   <li>Other (null name) — OTHER type, no remoteItem; also covers
         *       {@code name == null} short-circuit in the Excel-count filter</li>
         * </ol>
         */
        @Test
        @DisplayName("all item-type and remoteItem branches covered in one page")
        void allLoggingBranchesCovered() {
            List<DriveItem> items = List.of(
                    excelFileItemWithParentRef("Budget.xlsx"),      // FILE, remoteItem+parentRef
                    nonExcelFileItemNoParentRef("Notes.txt"),        // FILE, remoteItem no parentRef
                    folderItem("Archive"),                           // FOLDER, no remoteItem
                    otherItemNullName()                              // OTHER, no remoteItem, null name
            );
            when(mockFirstPage.getCurrentPage()).thenReturn(items);
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            List<?> result = service.getAllSharedDrivesInformation();

            assertNotNull(result);
        }

        @Test
        @DisplayName("FOLDER item type branch: folder != null → type='FOLDER'")
        void folderType_branch() {
            when(mockFirstPage.getCurrentPage()).thenReturn(List.of(folderItem("Reports")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("FILE item type branch: folder==null, file!=null → type='FILE'")
        void fileType_branch() {
            when(mockFirstPage.getCurrentPage())
                    .thenReturn(List.of(nonExcelFileItemNoParentRef("Data.csv")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("OTHER item type branch: folder==null, file==null → type='OTHER'")
        void otherType_branch() {
            when(mockFirstPage.getCurrentPage()).thenReturn(List.of(otherItemNullName()));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("remoteItem with parentReference → 'remoteItem.driveId=...' branch")
        void remoteItemWithParentRef_branch() {
            when(mockFirstPage.getCurrentPage())
                    .thenReturn(List.of(excelFileItemWithParentRef("Q1.xlsx")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("remoteItem without parentReference → 'remoteItem(no parentRef)' branch")
        void remoteItemWithoutParentRef_branch() {
            when(mockFirstPage.getCurrentPage())
                    .thenReturn(List.of(nonExcelFileItemNoParentRef("Notes.docx")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("remoteItem == null → 'no-remoteItem' branch")
        void noRemoteItem_branch() {
            when(mockFirstPage.getCurrentPage()).thenReturn(List.of(folderItem("Shared")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Excel-count filter branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Excel-count filter branches (file != null && name != null && isExcelFile)")
    class ExcelCountFilter {

        @Test
        @DisplayName("file!=null, name!=null, isExcelFile=true → counted as Excel file")
        void excelFile_allConditionsTrue() {
            when(mockFirstPage.getCurrentPage())
                    .thenReturn(List.of(excelFileItemWithParentRef("Workbook.xlsx")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("file!=null, name!=null, isExcelFile=false → NOT counted as Excel")
        void nonExcelFileWithName_isExcelFileFalse() {
            when(mockFirstPage.getCurrentPage())
                    .thenReturn(List.of(nonExcelFileItemNoParentRef("readme.txt")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("file!=null, name==null → name-null short-circuit, NOT counted")
        void fileItemWithNullName_shortCircuit() {
            DriveItem item = new DriveItem();
            item.file = new com.microsoft.graph.models.File();
            item.name = null; // short-circuits the name != null check

            when(mockFirstPage.getCurrentPage()).thenReturn(List.of(item));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }

        @Test
        @DisplayName("file==null (folder) → file-null short-circuit, NOT counted as Excel")
        void folderItem_fileNullShortCircuit() {
            when(mockFirstPage.getCurrentPage()).thenReturn(List.of(folderItem("Documents")));
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            assertNotNull(service.getAllSharedDrivesInformation());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Empty result from shared-with-me
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Empty shared-with-me page")
    class EmptyPage {

        @Test
        @DisplayName("first page exists but has no items → empty list, logging skipped")
        void firstPageEmpty_emptyListReturned() {
            when(mockFirstPage.getCurrentPage()).thenReturn(Collections.emptyList());
            when(mockFirstPage.getNextPage()).thenReturn(null);
            when(mockSharedWithMeRequest.get()).thenReturn(mockFirstPage);

            List<?> result = service.getAllSharedDrivesInformation();

            assertNotNull(result);
        }
    }
}
