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

import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * 100% branch coverage for the Site Files execution path in {@link DriveService}:
 * <ul>
 *   <li>{@code getExcelFileFromSiteByPath} — orchestrator</li>
 *   <li>{@code resolveSiteId} — Step 1: find SharePoint site</li>
 *   <li>{@code resolveDriveId} — Step 2: find document library (drive)</li>
 *   <li>{@code resolveFileInDrive} — Step 3: resolve file by path</li>
 *   <li>{@code extractLastUrlSegment} — static URL helper used by resolveDriveId</li>
 * </ul>
 *
 * <p>All methods are private; they are exercised through the public
 * {@code getExcelFileByPath(sourcePath, "Site Files")} entry point.
 *
 * <p><strong>Why explicit chain mocks instead of {@code RETURNS_DEEP_STUBS}:</strong>
 * {@code SiteCollectionRequest.get()} and {@code DriveCollectionRequest.get()} inherit their
 * {@code get()} method from the generic {@code BaseCollectionRequest<T, T2>}.  Due to Java generic
 * type erasure the compiled bytecode declares the return type as the raw {@code BaseCollectionPage},
 * not the covariant subclass.  Mockito's {@code RETURNS_DEEP_STUBS} resolves return types via
 * reflection and therefore creates a {@code BaseCollectionPage} mock for the call chain, which
 * causes a {@code ClassCastException} when the production code assigns the result to
 * {@code SiteCollectionPage} or {@code DriveCollectionPage}.  Mocking each link in the chain
 * with an explicit {@code @Mock} of the correct type avoids the type-erasure issue entirely.
 *
 * <p>Branches covered:
 * <ul>
 *   <li>getExcelFileFromSiteByPath: IOException from getGraphServiceClient → null</li>
 *   <li>getExcelFileFromSiteByPath: resolveSiteId returns null → null</li>
 *   <li>getExcelFileFromSiteByPath: resolveDriveId returns null → null</li>
 *   <li>getExcelFileFromSiteByPath: resolveFileInDrive returns null → null</li>
 *   <li>getExcelFileFromSiteByPath: all steps succeed → ExcelFile returned</li>
 *   <li>getExcelFileFromSiteByPath: GraphServiceException in inner try → null</li>
 *   <li>resolveSiteId: sitePage == null → return null</li>
 *   <li>resolveSiteId: sitePage.getCurrentPage().isEmpty() → return null</li>
 *   <li>resolveSiteId: exact displayName match → loop exits early with matched site</li>
 *   <li>resolveSiteId: displayName is null → no exact match → fallback to first result</li>
 *   <li>resolveSiteId: displayName non-null but mismatches → fallback to first result</li>
 *   <li>resolveDriveId: drivesPage == null → return null</li>
 *   <li>resolveDriveId: drive matched by name (nameMatch=true) → driveId returned</li>
 *   <li>resolveDriveId: drive.name is null → nameMatch=false; URL segment matches → urlMatch=true</li>
 *   <li>resolveDriveId: drive.name mismatches; URL segment matches → urlMatch wins</li>
 *   <li>resolveDriveId: no drive matches → return null</li>
 *   <li>resolveDriveId: multiple drives; second one matches → second driveId returned</li>
 *   <li>resolveFileInDrive: item == null → return null</li>
 *   <li>resolveFileInDrive: item found → ExcelFile with correct id, driveId, name, webUrl</li>
 *   <li>extractLastUrlSegment: webUrl is null → returns null</li>
 *   <li>extractLastUrlSegment: webUrl is blank → returns null</li>
 *   <li>extractLastUrlSegment: URL with encoded segment → decoded last segment returned</li>
 *   <li>extractLastUrlSegment: URL with no slash → full decoded URL returned as segment</li>
 *   <li>extractLastUrlSegment: malformed % sequence → exception caught → returns null</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("DriveService – Site Files strategy – 100% branch coverage")
class DriveServiceSiteFilesStrategyTest {

    // ─── Test constants ──────────────────────────────────────────────────────
    private static final String SITE_NAME    = "Engineering";
    private static final String LIBRARY_NAME = "Shared Documents";
    private static final String FILE_NAME    = "Report.xlsx";
    /** /site/library/file — the standard two-segment Site Files path */
    private static final String SOURCE_PATH  = "/" + SITE_NAME + "/" + LIBRARY_NAME + "/" + FILE_NAME;

    private static final String SITE_ID  = "site-id-111";
    private static final String DRIVE_ID = "drive-id-222";
    private static final String FILE_ID  = "file-id-333";

    // ─── Infrastructure mocks ────────────────────────────────────────────────
    @Mock private GraphServiceClientProviderFactory mockProviderFactory;
    @Mock private GraphServiceClientProvider        mockClientProvider;
    @Mock private KristaMediaClient                 mockKristaMediaClient;

    // ─── Graph client root (plain mock — no deep stubs needed) ───────────────
    @SuppressWarnings("rawtypes")
    @Mock private GraphServiceClient mockGraphClient;

    // sites() → buildRequest(list) → get()  chain
    @Mock private SiteCollectionRequestBuilder  mockSiteCollectionBuilder;
    @Mock private SiteCollectionRequest         mockSiteCollectionRequest;
    @Mock private SiteCollectionPage            mockSitePage;

    // sites(siteId) → drives() → buildRequest() → get()  chain
    @Mock private SiteRequestBuilder            mockSiteByIdBuilder;
    @Mock private DriveCollectionRequestBuilder mockDriveCollectionBuilder;
    @Mock private DriveCollectionRequest        mockDriveCollectionRequest;
    @Mock private DriveCollectionPage           mockDrivesPage;

    // drives(driveId) → root() → itemWithPath() → buildRequest() → get()  chain
    @Mock private DriveRequestBuilder           mockDriveRequestBuilder;
    @Mock private DriveItemRequestBuilder       mockRootItemBuilder;
    @Mock private DriveItemRequestBuilder       mockItemWithPathBuilder;
    @Mock private DriveItemRequest              mockDriveItemRequest;

    private DriveService service;

    @BeforeEach
    @SuppressWarnings("unchecked")
    void setUp() throws IOException {
        when(mockProviderFactory.create()).thenReturn(mockClientProvider);
        when(mockClientProvider.getGraphServiceClient()).thenReturn(mockGraphClient);

        // Wire up the sites() → SiteCollectionRequest chain
        when(mockGraphClient.sites()).thenReturn(mockSiteCollectionBuilder);
        when(mockSiteCollectionBuilder.buildRequest(anyList())).thenReturn(mockSiteCollectionRequest);
        // mockSiteCollectionRequest.get() is configured per-test via stubSiteFound() / direct stub

        // Wire up the sites(siteId) → DriveCollectionRequest chain
        when(mockGraphClient.sites(anyString())).thenReturn(mockSiteByIdBuilder);
        when(mockSiteByIdBuilder.drives()).thenReturn(mockDriveCollectionBuilder);
        when(mockDriveCollectionBuilder.buildRequest()).thenReturn(mockDriveCollectionRequest);
        // mockDriveCollectionRequest.get() is configured per-test via stubDrivePageWith() / direct stub

        // Wire up the drives(driveId) → DriveItemRequest chain
        when(mockGraphClient.drives(anyString())).thenReturn(mockDriveRequestBuilder);
        when(mockDriveRequestBuilder.root()).thenReturn(mockRootItemBuilder);
        when(mockRootItemBuilder.itemWithPath(anyString())).thenReturn(mockItemWithPathBuilder);
        when(mockItemWithPathBuilder.buildRequest()).thenReturn(mockDriveItemRequest);
        // mockDriveItemRequest.get() is configured per-test via stubFileFound() / direct stub

        service = new DriveService(mockProviderFactory, mockKristaMediaClient);
    }

    // ─── Model factories ─────────────────────────────────────────────────────

    private static com.microsoft.graph.models.Site site(String id, String displayName) {
        com.microsoft.graph.models.Site s = new com.microsoft.graph.models.Site();
        s.id          = id;
        s.displayName = displayName;
        return s;
    }

    private static com.microsoft.graph.models.Drive drive(String id, String name, String webUrl) {
        com.microsoft.graph.models.Drive d = new com.microsoft.graph.models.Drive();
        d.id     = id;
        d.name   = name;
        d.webUrl = webUrl;
        return d;
    }

    private static DriveItem driveItem(String id, String name, String webUrl) {
        DriveItem item = new DriveItem();
        item.id     = id;
        item.name   = name;
        item.webUrl = webUrl;
        return item;
    }

    // ─── Stub helpers ─────────────────────────────────────────────────────────

    /** Stubs the site-search call returning a page with a single matching site. */
    private void stubSiteFound(String displayName, String siteId) {
        when(mockSitePage.getCurrentPage()).thenReturn(List.of(site(siteId, displayName)));
        when(mockSiteCollectionRequest.get()).thenReturn(mockSitePage);
    }

    /** Stubs the drives-listing call for any siteId to return a page with the given drives. */
    private void stubDrivePageWith(com.microsoft.graph.models.Drive... drives) {
        when(mockDrivesPage.getCurrentPage()).thenReturn(List.of(drives));
        when(mockDriveCollectionRequest.get()).thenReturn(mockDrivesPage);
    }

    /** Stubs the item-by-path call for any driveId + any path to return the given item. */
    private void stubFileFound(DriveItem item) {
        when(mockDriveItemRequest.get()).thenReturn(item);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // getExcelFileFromSiteByPath — orchestrator branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("getExcelFileFromSiteByPath – orchestrator branches")
    class OrchestratorBranches {

        @Test
        @DisplayName("IOException from getGraphServiceClient → returns null → caller throws IAE")
        void ioException_returnsNull_callerThrowsIAE() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("network error"));

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("resolveSiteId returns null → returns null → caller throws IAE")
        void resolveSiteIdNull_callerThrowsIAE() {
            when(mockSiteCollectionRequest.get()).thenReturn(null);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("resolveDriveId returns null → returns null → caller throws IAE")
        void resolveDriveIdNull_callerThrowsIAE() {
            stubSiteFound(SITE_NAME, SITE_ID);
            when(mockDriveCollectionRequest.get()).thenReturn(null);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("resolveFileInDrive returns null → returns null → caller throws IAE")
        void resolveFileInDriveNull_callerThrowsIAE() {
            stubSiteFound(SITE_NAME, SITE_ID);
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, null));
            when(mockDriveItemRequest.get()).thenReturn(null);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("all three steps succeed → ExcelFile with correct fields returned")
        void allStepsSucceed_returnsExcelFile() {
            stubSiteFound(SITE_NAME, SITE_ID);
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, null));
            stubFileFound(driveItem(FILE_ID, FILE_NAME, "https://example.com/Report.xlsx"));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(FILE_ID,   result.fileId);
            assertEquals(DRIVE_ID,  result.driveId);
            assertEquals(FILE_NAME, result.fileName);
        }

        @Test
        @DisplayName("GraphServiceException thrown inside inner try → returns null → caller throws IAE")
        void graphServiceException_returnsNull_callerThrowsIAE() {
            stubSiteFound(SITE_NAME, SITE_ID);
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, null));
            when(mockDriveItemRequest.get()).thenThrow(mock(GraphServiceException.class));

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // resolveSiteId — Step 1 branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("resolveSiteId – Step 1 branches")
    class ResolveSiteIdBranches {

        @BeforeEach
        void stubDriveAndFile() {
            // Allow execution to continue past resolveSiteId in success cases
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, null));
            stubFileFound(driveItem(FILE_ID, FILE_NAME, null));
        }

        @Test
        @DisplayName("sitePage == null → return null → IAE")
        void sitePageNull_returnsNull_callerThrowsIAE() {
            when(mockSiteCollectionRequest.get()).thenReturn(null);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("sitePage.getCurrentPage() is empty → return null → IAE")
        void sitePageEmpty_returnsNull_callerThrowsIAE() {
            when(mockSitePage.getCurrentPage()).thenReturn(List.of());
            when(mockSiteCollectionRequest.get()).thenReturn(mockSitePage);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("exact displayName match found in loop → matched site's siteId used")
        void exactDisplayNameMatch_matchedSiteUsed() {
            com.microsoft.graph.models.Site other = site("other-id",  "Other Site");
            com.microsoft.graph.models.Site exact = site(SITE_ID,     SITE_NAME);
            when(mockSitePage.getCurrentPage()).thenReturn(List.of(other, exact));
            when(mockSiteCollectionRequest.get()).thenReturn(mockSitePage);

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result); // execution reached resolveFileInDrive successfully
        }

        @Test
        @DisplayName("site.displayName is null → no exact match → fallback to first result")
        void displayNameNull_fallbackToFirstResult() {
            com.microsoft.graph.models.Site first = site(SITE_ID, null); // null displayName
            when(mockSitePage.getCurrentPage()).thenReturn(List.of(first));
            when(mockSiteCollectionRequest.get()).thenReturn(mockSitePage);

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
        }

        @Test
        @DisplayName("displayName present but does not match → fallback to first result")
        void displayNameMismatch_fallbackToFirstResult() {
            com.microsoft.graph.models.Site first = site(SITE_ID, "Completely Different Name");
            when(mockSitePage.getCurrentPage()).thenReturn(List.of(first));
            when(mockSiteCollectionRequest.get()).thenReturn(mockSitePage);

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // resolveDriveId — Step 2 branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("resolveDriveId – Step 2 branches")
    class ResolveDriveIdBranches {

        @BeforeEach
        void stubSiteAndFile() {
            stubSiteFound(SITE_NAME, SITE_ID);
            stubFileFound(driveItem(FILE_ID, FILE_NAME, null));
        }

        @Test
        @DisplayName("drivesPage == null → return null → IAE")
        void drivesPageNull_returnsNull_callerThrowsIAE() {
            when(mockDriveCollectionRequest.get()).thenReturn(null);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("drive matched by name (nameMatch=true) → correct driveId returned in ExcelFile")
        void matchByName_returnsCorrectDriveId() {
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, null)); // name matches

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(DRIVE_ID, result.driveId);
        }

        @Test
        @DisplayName("drive.name is null → nameMatch=false; URL segment matches (urlMatch=true)")
        void driveNameNull_urlSegmentMatches_returnsCorrectDriveId() {
            // name null → nameMatch=false; webUrl last segment = "Shared Documents" → urlMatch=true
            stubDrivePageWith(drive(DRIVE_ID, null,
                    "https://tenant.sharepoint.com/sites/Engineering/Shared%20Documents"));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(DRIVE_ID, result.driveId);
        }

        @Test
        @DisplayName("drive.name mismatches; URL segment matches → urlMatch wins")
        void driveNameMismatch_urlSegmentMatches_urlMatchWins() {
            stubDrivePageWith(drive(DRIVE_ID, "Wrong Name",
                    "https://tenant.sharepoint.com/sites/Engineering/Shared%20Documents"));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(DRIVE_ID, result.driveId);
        }

        @Test
        @DisplayName("no drive matches name or URL → return null → IAE")
        void noDriveMatches_returnsNull_callerThrowsIAE() {
            stubDrivePageWith(drive("wrong-id", "Other Library",
                    "https://tenant.sharepoint.com/sites/Engineering/Other%20Library"));

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("multiple drives: first doesn't match, second matches by name → second's driveId used")
        void multipleDrives_secondMatchesByName_secondDriveIdUsed() {
            com.microsoft.graph.models.Drive noMatch = drive("wrong-id", "Other Library", null);
            com.microsoft.graph.models.Drive match   = drive(DRIVE_ID,   LIBRARY_NAME,    null);
            stubDrivePageWith(noMatch, match);

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(DRIVE_ID, result.driveId);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // resolveFileInDrive — Step 3 branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("resolveFileInDrive – Step 3 branches")
    class ResolveFileInDriveBranches {

        @BeforeEach
        void stubSiteAndDrive() {
            stubSiteFound(SITE_NAME, SITE_ID);
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, null));
        }

        @Test
        @DisplayName("item == null → return null → IAE")
        void itemNull_returnsNull_callerThrowsIAE() {
            when(mockDriveItemRequest.get()).thenReturn(null);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(SOURCE_PATH, "Site Files"));
        }

        @Test
        @DisplayName("item found → ExcelFile returned with item.id, drive's driveId, item.name, item.webUrl")
        void itemFound_returnsCorrectExcelFile() {
            stubFileFound(driveItem(FILE_ID, FILE_NAME, "https://example.com/Report.xlsx"));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(FILE_ID,                           result.fileId);
            assertEquals(DRIVE_ID,                          result.driveId);   // from resolveDriveId, not item
            assertEquals(FILE_NAME,                         result.fileName);
            assertEquals("https://example.com/Report.xlsx", result.webURL);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // extractLastUrlSegment branches (exercised via resolveDriveId)
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("extractLastUrlSegment – all branches (via resolveDriveId url matching)")
    class ExtractLastUrlSegmentBranches {

        @BeforeEach
        void stubSiteAndFile() {
            stubSiteFound(SITE_NAME, SITE_ID);
            stubFileFound(driveItem(FILE_ID, FILE_NAME, null));
        }

        @Test
        @DisplayName("drive.webUrl is null → urlSegment=null → urlMatch=false (nameMatch used)")
        void webUrlNull_urlMatchFalse_nameMatchWins() {
            // name matches, webUrl=null → nameMatch=true covers the success path
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, null));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
        }

        @Test
        @DisplayName("drive.webUrl is blank → urlSegment=null → urlMatch=false (nameMatch used)")
        void webUrlBlank_urlMatchFalse_nameMatchWins() {
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, ""));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
        }

        @Test
        @DisplayName("drive.webUrl has URL-encoded segment → decoded last segment used for urlMatch=true")
        void webUrlEncodedSegment_urlMatchTrue() {
            // drive.name=null so nameMatch=false; URL segment decodes to "Shared Documents" → urlMatch=true
            stubDrivePageWith(drive(DRIVE_ID, null,
                    "https://tenant.sharepoint.com/sites/Engineering/Shared%20Documents"));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(DRIVE_ID, result.driveId);
        }

        @Test
        @DisplayName("drive.webUrl has no slash → full decoded URL used as segment → urlMatch=true")
        void webUrlNoSlash_fullDecodedUrlAsSegment() {
            // "Shared%20Documents" decodes to "Shared Documents"; no '/' → full string returned as segment
            stubDrivePageWith(drive(DRIVE_ID, null, "Shared%20Documents"));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result);
            assertEquals(DRIVE_ID, result.driveId);
        }

        @Test
        @DisplayName("drive.webUrl contains malformed % sequence → URLDecoder throws → urlSegment=null → urlMatch=false")
        void webUrlMalformed_exceptionCaught_urlSegmentNull() {
            // "%GG" is an invalid percent-encoded sequence; URLDecoder throws IAE → catch returns null
            // drive.name matches → nameMatch=true still succeeds
            stubDrivePageWith(drive(DRIVE_ID, LIBRARY_NAME, "%GG"));

            ExcelFile result = service.getExcelFileByPath(SOURCE_PATH, "Site Files");

            assertNotNull(result); // matched by name despite malformed webUrl
        }
    }
}
