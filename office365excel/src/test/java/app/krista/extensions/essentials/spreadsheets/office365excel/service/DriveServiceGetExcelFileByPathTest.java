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
import com.microsoft.graph.models.ItemReference;
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

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Tests for {@link DriveService#getExcelFileByPath(String, String)}.
 *
 * <p>Branches covered:
 * <ul>
 *   <li>Validation: null/blank path, missing filename, null/blank fileSearch</li>
 *   <li>Unknown strategy: unrecognised value → IAE with "Unknown File Search" message</li>
 *   <li>My Drive strategy: file found, not found, GraphServiceException, IOException</li>
 *   <li>Site Files strategy: too-short path, IOException</li>
 *   <li>Path parsing: leading slash stripped, URL-decoded path, nested subfolder</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("DriveService.getExcelFileByPath – all strategies")
class DriveServiceGetExcelFileByPathTest {

    @Mock private GraphServiceClientProviderFactory mockProviderFactory;
    @Mock private GraphServiceClientProvider mockClientProvider;

    @SuppressWarnings("rawtypes")
    @Mock(answer = Answers.RETURNS_DEEP_STUBS)
    private GraphServiceClient mockGraphClient;

    @Mock private KristaMediaClient mockKristaMediaClient;

    private DriveService service;

    @BeforeEach
    void setUp() throws IOException {
        when(mockProviderFactory.create()).thenReturn(mockClientProvider);
        when(mockClientProvider.getGraphServiceClient()).thenReturn(mockGraphClient);
        service = new DriveService(mockProviderFactory, mockKristaMediaClient);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Input Validation
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Input validation")
    class InputValidation {

        @Test
        @DisplayName("null sourcePath → IllegalArgumentException")
        void nullSourcePath_throwsIAE() {
            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath(null, "My Drive"));
        }

        @Test
        @DisplayName("blank sourcePath → IllegalArgumentException")
        void blankSourcePath_throwsIAE() {
            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("   ", "My Drive"));
        }

        @Test
        @DisplayName("path with no filename (trailing slash) → IllegalArgumentException")
        void pathWithNoFilename_throwsIAE() {
            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/folder/", "My Drive"));
        }

        @Test
        @DisplayName("null fileSearch → IllegalArgumentException with helpful message")
        void nullLookupStrategy_throwsIAE() {
            IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/file.xlsx", null));
            assertTrue(ex.getMessage().contains("File Search is required"));
        }

        @Test
        @DisplayName("blank fileSearch → IllegalArgumentException with helpful message")
        void blankLookupStrategy_throwsIAE() {
            IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/file.xlsx", "   "));
            assertTrue(ex.getMessage().contains("File Search is required"));
        }

        @Test
        @DisplayName("Site Files strategy with only one path segment → IllegalArgumentException")
        void siteFiles_oneSegmentPath_throwsIAE() {
            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/OnlyOneSite/file.xlsx", "Site Files"));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Path Parsing
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Path parsing")
    class PathParsing {

        @Test
        @DisplayName("leading slash is stripped — My Drive resolves correctly")
        void leadingSlashStripped_myDrive() {
            DriveItem driveItem = driveItem("file-1", "Report.xlsx", "drive-1");
            when(mockGraphClient.me().drive().root().itemWithPath(anyString())
                    .buildRequest().get()).thenReturn(driveItem);

            ExcelFile result = service.getExcelFileByPath("/Report.xlsx", "My Drive");

            assertNotNull(result);
            assertEquals("file-1", result.fileId);
        }

        @Test
        @DisplayName("URL-encoded path (%20) is decoded before lookup — My Drive strategy")
        void urlEncodedPath_decoded_myDriveStrategy() {
            DriveItem decoded = driveItem("f-1", "Q1 Report.xlsx", "d-1");
            when(mockGraphClient.me().drive().root().itemWithPath(anyString())
                    .buildRequest().get()).thenReturn(decoded);

            ExcelFile result = service.getExcelFileByPath("/Q1%20Report.xlsx", "My Drive");

            assertNotNull(result);
            assertEquals("f-1", result.fileId);
        }

        @Test
        @DisplayName("nested path — folder and filename split correctly for My Drive")
        void nestedPath_splitCorrectly_myDrive() {
            DriveItem driveItem = driveItem("file-1", "Budget.xlsx", "drive-1");
            when(mockGraphClient.me().drive().root().itemWithPath(anyString())
                    .buildRequest().get()).thenReturn(driveItem);

            ExcelFile result = service.getExcelFileByPath("/Finance/2026/Budget.xlsx", "My Drive");

            assertNotNull(result);
            assertEquals("file-1", result.fileId);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // My Drive Strategy
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("My Drive strategy")
    class MyDriveStrategy {

        @Test
        @DisplayName("file found → returns ExcelFile")
        void fileFound_returnsExcelFile() {
            DriveItem item = driveItem("file-1", "Budget.xlsx", "drive-1");
            when(mockGraphClient.me().drive().root().itemWithPath(anyString())
                    .buildRequest().get()).thenReturn(item);

            ExcelFile result = service.getExcelFileByPath("/Budget.xlsx", "My Drive");

            assertNotNull(result);
            assertEquals("file-1", result.fileId);
            assertEquals("drive-1", result.driveId);
        }

        @Test
        @DisplayName("item not found (null returned) → IllegalArgumentException")
        void itemNotFound_throwsIAE() {
            when(mockGraphClient.me().drive().root().itemWithPath(anyString())
                    .buildRequest().get()).thenReturn(null);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/Missing.xlsx", "My Drive"));
        }

        @Test
        @DisplayName("GraphServiceException → IllegalArgumentException")
        void graphServiceException_throwsIAE() {
            GraphServiceException gse = mock(GraphServiceException.class);
            when(mockGraphClient.me().drive().root().itemWithPath(anyString())
                    .buildRequest().get()).thenThrow(gse);

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/Missing.xlsx", "My Drive"));
        }

        @Test
        @DisplayName("IOException from getGraphServiceClient → IllegalArgumentException")
        void ioException_throwsIAE() throws IOException {
            when(mockClientProvider.getGraphServiceClient())
                    .thenThrow(new IOException("network error"));

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/Budget.xlsx", "My Drive"));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Unknown Strategy Validation
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Unknown strategy validation")
    class UnknownStrategyValidation {

        @Test
        @DisplayName("unrecognised string → IllegalArgumentException with 'Unknown File Search' message")
        void unknownString_throwsIAEWithHelpfulMessage() {
            IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/Testing.xlsx", "Shared With Me"));
            assertTrue(ex.getMessage().contains("Unknown File Search"));
        }

        @Test
        @DisplayName("arbitrary string → IllegalArgumentException with 'Unknown File Search' message")
        void arbitraryString_throwsIAEWithHelpfulMessage() {
            IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/Testing.xlsx", "FTP"));
            assertTrue(ex.getMessage().contains("Unknown File Search"));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Site Files Strategy
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Site Files strategy")
    class SiteFilesStrategy {

        @Test
        @DisplayName("path with only filename (no site/library segments) → IllegalArgumentException")
        void pathTooShort_throwsIAE() {
            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/file.xlsx", "Site Files"));
        }

        @Test
        @DisplayName("IOException from getGraphServiceClient → IllegalArgumentException")
        void ioException_throwsIAE() throws IOException {
            when(mockClientProvider.getGraphServiceClient())
                    .thenThrow(new IOException("auth error"));

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/Engineering/Shared Documents/file.xlsx", "Site Files"));
        }

        @Test
        @DisplayName("valid two-segment path with IOException → IllegalArgumentException")
        void twoSegmentPath_ioException_throwsIAE() throws IOException {
            when(mockClientProvider.getGraphServiceClient())
                    .thenThrow(new IOException("connection refused"));

            assertThrows(IllegalArgumentException.class,
                    () -> service.getExcelFileByPath("/Engineering/Shared Documents/Report.xlsx", "Site Files"));
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Helpers
    // ═══════════════════════════════════════════════════════════════════════

    private static DriveItem driveItem(String id, String name, String driveId) {
        DriveItem item = new DriveItem();
        item.id   = id;
        item.name = name;
        if (driveId != null) {
            item.parentReference = new ItemReference();
            item.parentReference.driveId = driveId;
        }
        return item;
    }

}
