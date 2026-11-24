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

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import app.krista.model.base.File;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

/**
 * Comprehensive test suite for FileService with 100% code coverage.
 * Tests CSV file reading and row extraction functionality.
 */
@ExtendWith(MockitoExtension.class)
class FileServiceTest {

    @Mock
    private KristaMediaClient kristaMediaClient;

    private FileService fileService;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        fileService = new FileService(kristaMediaClient);
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_initializesWithKristaMediaClient() {
        FileService service = new FileService(kristaMediaClient);
        assertNotNull(service);
    }

    // ==================== getRows Tests - Success Cases ====================

    @Test
    void getRows_withSingleCsvFile_returnsAllRows() throws IOException {
        // Create a temporary CSV file
        Path csvFile = tempDir.resolve("test.csv");
        List<String> lines = List.of("header1,header2,header3", "value1,value2,value3", "value4,value5,value6");
        Files.write(csvFile, lines, StandardCharsets.UTF_8);

        // Mock the KristaMediaClient
        File mockFile = mock(File.class);
        when(kristaMediaClient.toJavaFile(mockFile)).thenReturn(csvFile.toFile());

        List<File> csvFiles = List.of(mockFile);

        List<String> result = fileService.getRows(csvFiles);

        assertNotNull(result);
        assertEquals(3, result.size());
        assertEquals("header1,header2,header3", result.get(0));
        assertEquals("value1,value2,value3", result.get(1));
        assertEquals("value4,value5,value6", result.get(2));
        verify(kristaMediaClient).toJavaFile(mockFile);
    }

    @Test
    void getRows_withMultipleCsvFiles_combinesAllRows() throws IOException {
        // Create two temporary CSV files
        Path csvFile1 = tempDir.resolve("test1.csv");
        Path csvFile2 = tempDir.resolve("test2.csv");
        Files.write(csvFile1, List.of("row1", "row2"), StandardCharsets.UTF_8);
        Files.write(csvFile2, List.of("row3", "row4"), StandardCharsets.UTF_8);

        File mockFile1 = mock(File.class);
        File mockFile2 = mock(File.class);
        when(kristaMediaClient.toJavaFile(mockFile1)).thenReturn(csvFile1.toFile());
        when(kristaMediaClient.toJavaFile(mockFile2)).thenReturn(csvFile2.toFile());

        List<File> csvFiles = List.of(mockFile1, mockFile2);

        List<String> result = fileService.getRows(csvFiles);

        assertNotNull(result);
        assertEquals(4, result.size());
        assertEquals("row1", result.get(0));
        assertEquals("row2", result.get(1));
        assertEquals("row3", result.get(2));
        assertEquals("row4", result.get(3));
    }

    @Test
    void getRows_withEmptyCsvFile_returnsEmptyList() throws IOException {
        Path csvFile = tempDir.resolve("empty.csv");
        Files.write(csvFile, List.of(), StandardCharsets.UTF_8);

        File mockFile = mock(File.class);
        when(kristaMediaClient.toJavaFile(mockFile)).thenReturn(csvFile.toFile());

        List<File> csvFiles = List.of(mockFile);

        List<String> result = fileService.getRows(csvFiles);

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void getRows_withEmptyFileList_returnsEmptyList() throws IOException {
        List<File> csvFiles = new ArrayList<>();

        List<String> result = fileService.getRows(csvFiles);

        assertNotNull(result);
        assertTrue(result.isEmpty());
        verify(kristaMediaClient, never()).toJavaFile(any());
    }

    // ==================== getRows Tests - Special Characters ====================

    @Test
    void getRows_withUtf8Characters_preservesEncoding() throws IOException {
        Path csvFile = tempDir.resolve("utf8.csv");
        List<String> lines = List.of("名前,年齢", "太郎,25", "花子,30");
        Files.write(csvFile, lines, StandardCharsets.UTF_8);

        File mockFile = mock(File.class);
        when(kristaMediaClient.toJavaFile(mockFile)).thenReturn(csvFile.toFile());

        List<File> csvFiles = List.of(mockFile);

        List<String> result = fileService.getRows(csvFiles);

        assertNotNull(result);
        assertEquals(3, result.size());
        assertEquals("名前,年齢", result.get(0));
        assertEquals("太郎,25", result.get(1));
        assertEquals("花子,30", result.get(2));
    }

    @Test
    void getRows_withSpecialCharacters_preservesContent() throws IOException {
        Path csvFile = tempDir.resolve("special.csv");
        List<String> lines = List.of("col1,col2", "value with \"quotes\"", "value,with,commas");
        Files.write(csvFile, lines, StandardCharsets.UTF_8);

        File mockFile = mock(File.class);
        when(kristaMediaClient.toJavaFile(mockFile)).thenReturn(csvFile.toFile());

        List<File> csvFiles = List.of(mockFile);

        List<String> result = fileService.getRows(csvFiles);

        assertNotNull(result);
        assertEquals(3, result.size());
        assertTrue(result.get(1).contains("quotes"));
    }

    // ==================== getRows Tests - Error Cases ====================

    @Test
    void getRows_whenFileNotFound_throwsIOException() throws IOException {
        File mockFile = mock(File.class);
        when(kristaMediaClient.toJavaFile(mockFile)).thenReturn(new java.io.File("/nonexistent/file.csv"));

        List<File> csvFiles = List.of(mockFile);

        assertThrows(IOException.class, () -> fileService.getRows(csvFiles));
    }
}

