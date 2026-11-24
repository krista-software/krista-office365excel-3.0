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
import junit.framework.TestCase;
import org.junit.Assert;
import org.junit.Test;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

public class DriveServiceTest extends TestCase {

    DriveService driveService = Mockito.mock(DriveService.class);

    @Test()
    public void testGetPdfFile() throws IOException {
        Assert.assertNull(driveService.getPdfFile(null, null));
    }

    @Test()
    public void testGetPdfFileForGivenExcel() {
        Assert.assertEquals(new HashMap<>(), driveService.getPdfFileForGivenExcel(null, null));
    }

    @Test()
    public void testGetSharedExcelFilesInformation() {
        driveService.getSharedExcelFilesInformation("abcd");
    }

    // ── new-signature smoke tests (mock-based, matching existing style) ──

    @Test()
    public void testSearchExcelFilesInDrive_withXlsxFormat() {
        driveService.searchExcelFilesInDrive("drive-id", ".xlsx");
    }

    @Test()
    public void testSearchExcelFilesInDrive_withCsvFormat() {
        driveService.searchExcelFilesInDrive("drive-id", ".csv");
    }

    @Test()
    public void testSearchExcelFilesInDrive_withXlsFormat() {
        driveService.searchExcelFilesInDrive("drive-id", ".xls");
    }

    @Test()
    public void testSearchExcelFilesInDrive_withTsvFormat() {
        driveService.searchExcelFilesInDrive("drive-id", ".tsv");
    }

    @Test()
    public void testSearchExcelFilesInDrive_withNullFormat() {
        driveService.searchExcelFilesInDrive("drive-id", null);
    }

    @Test()
    public void testSearchExcelFilesInDrive_withBlankFormat() {
        driveService.searchExcelFilesInDrive("drive-id", "   ");
    }

    // ── unit tests for extension-selection logic ──────────────────────

    @ExtendWith(MockitoExtension.class)
    @MockitoSettings(strictness = Strictness.LENIENT)
    @DisplayName("DriveService.searchExcelFilesInDrive – extension selection and error handling")
    public static class SearchExcelFilesInDriveUnitTest {

        @Mock
        private GraphServiceClientProviderFactory mockProviderFactory;

        @Mock
        private GraphServiceClientProvider mockClientProvider;

        @Mock
        private KristaMediaClient mockKristaMediaClient;

        @org.junit.jupiter.api.BeforeEach
        void setUp() {
            when(mockProviderFactory.create()).thenReturn(mockClientProvider);
        }

        @org.junit.jupiter.api.Test
        @DisplayName("returns empty list when getGraphServiceClient throws IOException")
        void returnsEmptyList_whenGraphClientThrowsIOException() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("auth failure"));

            DriveService service = new DriveService(mockProviderFactory, mockKristaMediaClient);
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-id", ".xlsx");

            assertEquals(Collections.emptyList(), result);
        }

        @org.junit.jupiter.api.Test
        @DisplayName("returns empty list with null fileFormat when getGraphServiceClient throws IOException")
        void returnsEmptyList_withNullFormat_whenGraphClientThrowsIOException() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("auth failure"));

            DriveService service = new DriveService(mockProviderFactory, mockKristaMediaClient);
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-id", null);

            assertEquals(Collections.emptyList(), result);
        }

        @org.junit.jupiter.api.Test
        @DisplayName("returns empty list with blank fileFormat when getGraphServiceClient throws IOException")
        void returnsEmptyList_withBlankFormat_whenGraphClientThrowsIOException() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("network error"));

            DriveService service = new DriveService(mockProviderFactory, mockKristaMediaClient);
            List<ExcelFile> result = service.searchExcelFilesInDrive("drive-id", "  ");

            assertEquals(Collections.emptyList(), result);
        }

        @org.junit.jupiter.api.Test
        @DisplayName("getGraphServiceClient called exactly once regardless of fileFormat value")
        void graphClientFetchedOnce_forSpecificFormat() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("test"));

            DriveService service = new DriveService(mockProviderFactory, mockKristaMediaClient);
            service.searchExcelFilesInDrive("drive-id", ".csv");

            verify(mockClientProvider, times(1)).getGraphServiceClient();
        }

        @org.junit.jupiter.api.Test
        @DisplayName("getGraphServiceClient called exactly once when fileFormat is null")
        void graphClientFetchedOnce_forNullFormat() throws IOException {
            when(mockClientProvider.getGraphServiceClient()).thenThrow(new IOException("test"));

            DriveService service = new DriveService(mockProviderFactory, mockKristaMediaClient);
            service.searchExcelFilesInDrive("drive-id", null);

            verify(mockClientProvider, times(1)).getGraphServiceClient();
        }
    }
}
