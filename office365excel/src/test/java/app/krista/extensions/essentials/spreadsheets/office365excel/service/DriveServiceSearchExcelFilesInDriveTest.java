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
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.DriveItemSearchParameterSet;
import com.microsoft.graph.requests.DriveItemSearchCollectionPage;
import com.microsoft.graph.requests.DriveItemSearchCollectionRequest;
import com.microsoft.graph.requests.DriveItemSearchCollectionRequestBuilder;
import com.microsoft.graph.requests.GraphServiceClient;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Answers;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;
import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * 100% branch coverage for {@link DriveService#searchExcelFilesInDrive(String, String)}.
 *
 * <p>Branches covered:
 * <ol>
 *   <li>fileFormat not-null/non-blank  → single-element extension array (ternary true)</li>
 *   <li>fileFormat null or blank       → all-four extension array (ternary false)</li>
 *   <li>IOException from getGraphServiceClient → return empty list</li>
 *   <li>page == null                   → while-loop not entered</li>
 *   <li>item.id != null                → added to uniqueItems</li>
 *   <li>item.id == null                → skipped (not added)</li>
 *   <li>nextPage == null               → break out of while-loop</li>
 *   <li>nextPage != null               → fetch second page</li>
 *   <li>GraphServiceException          → logged, loop continues</li>
 *   <li>Deduplication                  → putIfAbsent keeps only first</li>
 * </ol>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("DriveService.searchExcelFilesInDrive – 100% branch coverage")
class DriveServiceSearchExcelFilesInDriveTest {

    // ── dependencies ──────────────────────────────────────────────────────
    @Mock private GraphServiceClientProviderFactory mockProviderFactory;
    @Mock private GraphServiceClientProvider mockClientProvider;

    @SuppressWarnings("rawtypes")
    @Mock(answer = Answers.RETURNS_DEEP_STUBS)
    private GraphServiceClient mockGraphClient;

    @Mock private KristaMediaClient mockKristaMediaClient;

    // ── page / pagination mocks ───────────────────────────────────────────
    @Mock private DriveItemSearchCollectionPage mockPage;

    @Mock
    private DriveItemSearchCollectionRequestBuilder mockNextPageBuilder;

    private DriveService service;

    @BeforeEach
    void setUp() throws IOException {
        when(mockProviderFactory.create()).thenReturn(mockClientProvider);
        when(mockClientProvider.getGraphServiceClient()).thenReturn(mockGraphClient);
        // Default: every Graph search call returns null (empty page)
        when(mockGraphClient.me().drives(anyString()).root()
                .search(any(DriveItemSearchParameterSet.class))
                .buildRequest().top(anyInt()).get())
                .thenReturn(null);
        service = new DriveService(mockProviderFactory, mockKristaMediaClient);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Branch C — IOException from getGraphServiceClient
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("IOException from getGraphServiceClient")
    class IOExceptionBranch {

        @Test
        @DisplayName("specific format → IOException returns empty list")
        void specificFormat_ioException_returnsEmptyList() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("auth error"));
            assertEquals(Collections.emptyList(), service.searchExcelFilesInDrive("drive-1", ".xlsx"));
        }

        @Test
        @DisplayName("null format → IOException returns empty list")
        void nullFormat_ioException_returnsEmptyList() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("auth error"));
            assertEquals(Collections.emptyList(), service.searchExcelFilesInDrive("drive-1", null));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Branches A/B — Extension array selection (ternary)
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Extension selection (ternary)")
    class ExtensionSelection {

        @Test
        @DisplayName("non-null non-blank format (.xlsx) → ternary true branch, null page → empty result")
        void nonNullFormat_xlsx_ternaryTrue_nullPage() {
            // default stub already returns null → while-loop not entered
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");
            assertNotNull(result);
            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("non-null non-blank format (.xls) → ternary true branch")
        void nonNullFormat_xls_ternaryTrue() {
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xls");
            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("non-null non-blank format (.csv) → ternary true branch")
        void nonNullFormat_csv_ternaryTrue() {
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".csv");
            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("non-null non-blank format (.tsv) → ternary true branch")
        void nonNullFormat_tsv_ternaryTrue() {
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".tsv");
            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("null fileFormat → ternary false branch, all 4 extensions, no pages → empty")
        void nullFormat_ternaryFalse_allExtensions_nullPages() {
            // default stub returns null for all 4 calls
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", null);
            assertNotNull(result);
            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("blank fileFormat → ternary false branch (isBlank check)")
        void blankFormat_ternaryFalse() {
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", "   ");
            assertNotNull(result);
            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("null format → items from all 4 extensions collected, 4 distinct files returned")
        void nullFormat_allFourExtensionsContributeItems() {
            DriveItemSearchCollectionPage p1 = pageWith(item("id-1", "A.xlsx"));
            DriveItemSearchCollectionPage p2 = pageWith(item("id-2", "B.xls"));
            DriveItemSearchCollectionPage p3 = pageWith(item("id-3", "C.csv"));
            DriveItemSearchCollectionPage p4 = pageWith(item("id-4", "D.tsv"));

            when(mockGraphClient.me().drives(anyString()).root()
                    .search(any(DriveItemSearchParameterSet.class))
                    .buildRequest().top(anyInt()).get())
                    .thenReturn(p1, p2, p3, p4);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", null);

            assertEquals(4, result.size());
            assertTrue(result.stream().anyMatch(f -> "id-1".equals(f.fileId)));
            assertTrue(result.stream().anyMatch(f -> "id-2".equals(f.fileId)));
            assertTrue(result.stream().anyMatch(f -> "id-3".equals(f.fileId)));
            assertTrue(result.stream().anyMatch(f -> "id-4".equals(f.fileId)));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Branch E2 — page == null (while-loop not entered)
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Null page (while-loop not entered)")
    class NullPageBranch {

        @Test
        @DisplayName("search returns null → no items collected, returns empty list")
        void searchReturnsNull_returnsEmptyList() {
            // default stub already returns null
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");
            assertTrue(result.isEmpty());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Branches F/G — item.id not null (added) and item.id null (skipped)
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Item id null / non-null")
    class ItemIdBranch {

        @Test
        @DisplayName("item.id not null → added to uniqueItems, returned in result")
        void itemIdNotNull_addedToResult() {
            DriveItem excelItem = item("file-1", "Report.xlsx");
            when(mockPage.getCurrentPage()).thenReturn(List.of(excelItem));
            when(mockPage.getNextPage()).thenReturn(null);
            stubSearchPage(mockPage);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");

            assertEquals(1, result.size());
            assertEquals("file-1", result.get(0).fileId);
            assertEquals("Report.xlsx", result.get(0).fileName);
        }

        @Test
        @DisplayName("item.id null → skipped, not added to uniqueItems, returns empty")
        void itemIdNull_skippedNotAddedToResult() {
            DriveItem nullIdItem = item(null, "Report.xlsx"); // id == null → branch H
            when(mockPage.getCurrentPage()).thenReturn(List.of(nullIdItem));
            when(mockPage.getNextPage()).thenReturn(null);
            stubSearchPage(mockPage);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");

            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("page has mixed items: one with null id, one with non-null id → only non-null returned")
        void mixedItems_onlyNonNullIdReturned() {
            DriveItem nullIdItem  = item(null,     "Ghost.xlsx");
            DriveItem validItem   = item("file-1", "Report.xlsx");
            when(mockPage.getCurrentPage()).thenReturn(List.of(nullIdItem, validItem));
            when(mockPage.getNextPage()).thenReturn(null);
            stubSearchPage(mockPage);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");

            assertEquals(1, result.size());
            assertEquals("file-1", result.get(0).fileId);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Branches I/J/K — pagination: next page present and next page absent
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Pagination")
    class PaginationBranch {

        @Test
        @DisplayName("nextPage == null → break out of while-loop (branch K)")
        void noNextPage_breaksLoop() {
            DriveItem it = item("file-1", "Data.csv");
            when(mockPage.getCurrentPage()).thenReturn(List.of(it));
            when(mockPage.getNextPage()).thenReturn(null); // → break
            stubSearchPage(mockPage);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".csv");

            assertEquals(1, result.size());
            assertEquals("file-1", result.get(0).fileId);
        }

        @Test
        @DisplayName("nextPage != null → second page fetched (branch J), items from both pages returned")
        void hasNextPage_secondPageFetched_itemsFromBothPages() {
            DriveItem item1 = item("file-1", "Report.xlsx");
            DriveItem item2 = item("file-2", "Budget.xlsx");

            DriveItemSearchCollectionPage page2 = mock(DriveItemSearchCollectionPage.class);
            when(page2.getCurrentPage()).thenReturn(List.of(item2));
            when(page2.getNextPage()).thenReturn(null); // second page has no further pages

            DriveItemSearchCollectionRequest nextPageReq =
                    mock(DriveItemSearchCollectionRequest.class);
            doReturn(page2).when(nextPageReq).get();
            doReturn(nextPageReq).when(mockNextPageBuilder).buildRequest();

            when(mockPage.getCurrentPage()).thenReturn(List.of(item1));
            when(mockPage.getNextPage()).thenReturn(mockNextPageBuilder); // → branch J
            stubSearchPage(mockPage);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");

            assertEquals(2, result.size());
            assertTrue(result.stream().anyMatch(f -> "file-1".equals(f.fileId)));
            assertTrue(result.stream().anyMatch(f -> "file-2".equals(f.fileId)));
        }

        @Test
        @DisplayName("three pages: page1 → page2 → page3 (null), all items collected")
        void threePagesOfResults_allItemsCollected() {
            DriveItem it1 = item("f-1", "A.xlsx");
            DriveItem it2 = item("f-2", "B.xlsx");
            DriveItem it3 = item("f-3", "C.xlsx");

            DriveItemSearchCollectionPage page2 = mock(DriveItemSearchCollectionPage.class);
            DriveItemSearchCollectionPage page3 = mock(DriveItemSearchCollectionPage.class);

            DriveItemSearchCollectionRequestBuilder nextBuilder2 =
                    mock(DriveItemSearchCollectionRequestBuilder.class);
            DriveItemSearchCollectionRequest req1 = mock(DriveItemSearchCollectionRequest.class);
            DriveItemSearchCollectionRequest req2 = mock(DriveItemSearchCollectionRequest.class);

            when(page3.getCurrentPage()).thenReturn(List.of(it3));
            when(page3.getNextPage()).thenReturn(null);

            when(page2.getCurrentPage()).thenReturn(List.of(it2));
            when(page2.getNextPage()).thenReturn(nextBuilder2);
            doReturn(req2).when(nextBuilder2).buildRequest();
            doReturn(page3).when(req2).get();

            when(mockPage.getCurrentPage()).thenReturn(List.of(it1));
            when(mockPage.getNextPage()).thenReturn(mockNextPageBuilder);
            doReturn(req1).when(mockNextPageBuilder).buildRequest();
            doReturn(page2).when(req1).get();
            stubSearchPage(mockPage);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");

            assertEquals(3, result.size());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Branch L — GraphServiceException catch block
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("GraphServiceException catch block")
    class GraphServiceExceptionBranch {

        @Test
        @DisplayName("exception thrown → caught, returns empty list (specific format)")
        void graphServiceException_specificFormat_returnsEmpty() {
            GraphServiceException gse = mock(GraphServiceException.class);
            when(mockGraphClient.me().drives(anyString()).root()
                    .search(any()).buildRequest().top(anyInt()).get())
                    .thenThrow(gse);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");

            assertTrue(result.isEmpty());
        }

        @Test
        @DisplayName("null format: exception on first extension (.xlsx), remaining 3 succeed → items from rest returned")
        void nullFormat_exceptionOnFirstExtension_continuesWithRemainingThree() {
            GraphServiceException gse = mock(GraphServiceException.class);

            DriveItemSearchCollectionPage p2 = pageWith(item("id-2", "B.xls"));
            DriveItemSearchCollectionPage p3 = pageWith(item("id-3", "C.csv"));
            DriveItemSearchCollectionPage p4 = pageWith(item("id-4", "D.tsv"));

            when(mockGraphClient.me().drives(anyString()).root()
                    .search(any()).buildRequest().top(anyInt()).get())
                    .thenThrow(gse)    // .xlsx → exception (branch L)
                    .thenReturn(p2)    // .xls  → success
                    .thenReturn(p3)    // .csv  → success
                    .thenReturn(p4);   // .tsv  → success

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", null);

            assertEquals(3, result.size());
        }

        @Test
        @DisplayName("null format: all 4 extensions throw exception → returns empty list")
        void nullFormat_allExtensionsThrow_returnsEmpty() {
            GraphServiceException gse = mock(GraphServiceException.class);
            when(mockGraphClient.me().drives(anyString()).root()
                    .search(any()).buildRequest().top(anyInt()).get())
                    .thenThrow(gse);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", null);

            assertTrue(result.isEmpty());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Deduplication — putIfAbsent
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Deduplication (putIfAbsent)")
    class DeduplicationBranch {

        @Test
        @DisplayName("same item id on two pages → putIfAbsent keeps only the first occurrence")
        void sameIdAcrossPages_onlyKeptOnce() {
            DriveItem item = item("file-1", "Report.xlsx");

            DriveItemSearchCollectionPage page2 = mock(DriveItemSearchCollectionPage.class);
            when(page2.getCurrentPage()).thenReturn(List.of(item)); // same id
            when(page2.getNextPage()).thenReturn(null);

            DriveItemSearchCollectionRequest nextPageReq =
                    mock(DriveItemSearchCollectionRequest.class);
            doReturn(page2).when(nextPageReq).get();
            doReturn(nextPageReq).when(mockNextPageBuilder).buildRequest();

            when(mockPage.getCurrentPage()).thenReturn(List.of(item)); // first occurrence
            when(mockPage.getNextPage()).thenReturn(mockNextPageBuilder);
            stubSearchPage(mockPage);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", ".xlsx");

            assertEquals(1, result.size()); // deduplicated
        }

        @Test
        @DisplayName("null format: same id returned by .xlsx and .xls searches → kept only once")
        void nullFormat_sameIdAcrossExtensions_deduplicatedInFinalResult() {
            DriveItem sharedItem = item("dup-id", "Shared.xlsx");
            DriveItem uniqueItem  = item("uni-id", "Unique.xls");

            DriveItemSearchCollectionPage p1 = pageWith(sharedItem);
            DriveItemSearchCollectionPage p2 = pageWith(sharedItem, uniqueItem); // dup-id seen again
            // p3 and p4 return null
            when(mockGraphClient.me().drives(anyString()).root()
                    .search(any()).buildRequest().top(anyInt()).get())
                    .thenReturn(p1, p2, null, null);

            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-1", null);

            // dup-id deduplicated → 2 unique files
            assertEquals(2, result.size());
            assertTrue(result.stream().anyMatch(f -> "dup-id".equals(f.fileId)));
            assertTrue(result.stream().anyMatch(f -> "uni-id".equals(f.fileId)));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Helpers
    // ═══════════════════════════════════════════════════════════════════════

    /** Stubs the search chain to return the given page. */
    private void stubSearchPage(DriveItemSearchCollectionPage page) {
        when(mockGraphClient.me().drives(anyString()).root()
                .search(any(DriveItemSearchParameterSet.class))
                .buildRequest().top(anyInt()).get())
                .thenReturn(page);
    }

    /** Creates a DriveItem with the given id and name. */
    private static DriveItem item(String id, String name) {
        DriveItem d = new DriveItem();
        d.id = id;
        d.name = name;
        return d;
    }

    /** Creates a single-page mock containing the given items, with no next page. */
    private static DriveItemSearchCollectionPage pageWith(DriveItem... items) {
        DriveItemSearchCollectionPage page = mock(DriveItemSearchCollectionPage.class);
        when(page.getCurrentPage()).thenReturn(List.of(items));
        when(page.getNextPage()).thenReturn(null);
        return page;
    }
}
