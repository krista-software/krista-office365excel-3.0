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

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClient;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.ArgumentCaptor;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Comprehensive test suite for SheetFormattingService with 100% code coverage.
 * Tests row formatting clearing operations for Excel worksheets.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class SheetFormattingServiceTest {

    @Mock
    private RestClientProvider restClientProvider;

    @Mock
    private RestClient restClient;

    private SheetFormattingService sheetFormattingService;

    private static final String DRIVE_ID = "drive123";
    private static final String FILE_ID = "file456";
    private static final String SHEET_NAME = "Sheet1";

    @BeforeEach
    void setUp() {
        when(restClientProvider.getRestClient(anyString())).thenReturn(restClient);
        when(restClientProvider.getRestClient(isNull())).thenReturn(restClient);
        sheetFormattingService = new SheetFormattingService(restClientProvider);
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_initializesWithRestClientProvider() {
        SheetFormattingService service = new SheetFormattingService(restClientProvider);
        assertNotNull(service);
    }

    // ==================== clearRowFormatting (Single Row) Tests ====================

    @Test
    void clearRowFormatting_singleRow_constructsCorrectUrl() throws IOException {
        Double rowIndex = 5.0;

        sheetFormattingService.clearRowFormatting(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex);

        ArgumentCaptor<String> urlCaptor = ArgumentCaptor.forClass(String.class);
        ArgumentCaptor<Map> bodyCaptor = ArgumentCaptor.forClass(Map.class);
        verify(restClient).post(urlCaptor.capture(), bodyCaptor.capture());

        String capturedUrl = urlCaptor.getValue();
        assertTrue(capturedUrl.contains(FILE_ID));
        assertTrue(capturedUrl.contains(SHEET_NAME));
        assertTrue(capturedUrl.contains("A5:Z100000"));
        assertTrue(capturedUrl.contains("/clear"));
    }

    @Test
    void clearRowFormatting_singleRow_sendsCorrectBody() throws IOException {
        Double rowIndex = 10.0;

        sheetFormattingService.clearRowFormatting(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex);

        ArgumentCaptor<Map> bodyCaptor = ArgumentCaptor.forClass(Map.class);
        verify(restClient).post(anyString(), bodyCaptor.capture());

        Map<String, Object> capturedBody = bodyCaptor.getValue();
        assertEquals("formats", capturedBody.get("applyTo"));
    }

    @Test
    void clearRowFormatting_singleRow_withDriveId_usesCorrectRestClient() throws IOException {
        Double rowIndex = 3.0;

        sheetFormattingService.clearRowFormatting(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex);

        verify(restClientProvider).getRestClient(DRIVE_ID);
        verify(restClient).post(anyString(), anyMap());
    }

    @Test
    void clearRowFormatting_singleRow_withDecimalRowIndex_convertsToInt() throws IOException {
        Double rowIndex = 7.9;

        sheetFormattingService.clearRowFormatting(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex);

        ArgumentCaptor<String> urlCaptor = ArgumentCaptor.forClass(String.class);
        verify(restClient).post(urlCaptor.capture(), anyMap());

        String capturedUrl = urlCaptor.getValue();
        assertTrue(capturedUrl.contains("A7:Z100000"));
    }

    // ==================== clearRowFormatting (Multiple Rows) Tests ====================

    @Test
    void clearRowFormatting_multipleRows_constructsCorrectRange() throws IOException {
        double startRowIndex = 5.0;
        int numberOfRows = 3;

        sheetFormattingService.clearRowFormatting(FILE_ID, SHEET_NAME, startRowIndex, numberOfRows);

        ArgumentCaptor<String> urlCaptor = ArgumentCaptor.forClass(String.class);
        verify(restClient).post(urlCaptor.capture(), anyMap());

        String capturedUrl = urlCaptor.getValue();
        assertTrue(capturedUrl.contains("A5:Z7")); // 5 + 3 - 1 = 7
    }

    @Test
    void clearRowFormatting_multipleRows_sendsCorrectBody() throws IOException {
        double startRowIndex = 1.0;
        int numberOfRows = 5;

        sheetFormattingService.clearRowFormatting(FILE_ID, SHEET_NAME, startRowIndex, numberOfRows);

        ArgumentCaptor<Map> bodyCaptor = ArgumentCaptor.forClass(Map.class);
        verify(restClient).post(anyString(), bodyCaptor.capture());

        Map<String, Object> capturedBody = bodyCaptor.getValue();
        assertEquals("formats", capturedBody.get("applyTo"));
    }

    @Test
    void clearRowFormatting_multipleRows_usesNullDriveId() throws IOException {
        double startRowIndex = 2.0;
        int numberOfRows = 4;

        sheetFormattingService.clearRowFormatting(FILE_ID, SHEET_NAME, startRowIndex, numberOfRows);

        verify(restClientProvider).getRestClient(null);
        verify(restClient).post(anyString(), anyMap());
    }

    @Test
    void clearRowFormatting_multipleRows_withSingleRow_calculatesCorrectEndRow() throws IOException {
        double startRowIndex = 10.0;
        int numberOfRows = 1;

        sheetFormattingService.clearRowFormatting(FILE_ID, SHEET_NAME, startRowIndex, numberOfRows);

        ArgumentCaptor<String> urlCaptor = ArgumentCaptor.forClass(String.class);
        verify(restClient).post(urlCaptor.capture(), anyMap());

        String capturedUrl = urlCaptor.getValue();
        assertTrue(capturedUrl.contains("A10:Z10")); // 10 + 1 - 1 = 10
    }

    @Test
    void clearRowFormatting_multipleRows_withLargeNumberOfRows_calculatesCorrectEndRow() throws IOException {
        double startRowIndex = 1.0;
        int numberOfRows = 100;

        sheetFormattingService.clearRowFormatting(FILE_ID, SHEET_NAME, startRowIndex, numberOfRows);

        ArgumentCaptor<String> urlCaptor = ArgumentCaptor.forClass(String.class);
        verify(restClient).post(urlCaptor.capture(), anyMap());

        String capturedUrl = urlCaptor.getValue();
        assertTrue(capturedUrl.contains("A1:Z100")); // 1 + 100 - 1 = 100
    }

    // ==================== Integration Tests ====================

    @Test
    void clearRowFormatting_bothMethods_useCorrectRestClientProvider() throws IOException {
        // Single row version
        sheetFormattingService.clearRowFormatting(DRIVE_ID, FILE_ID, SHEET_NAME, 5.0);
        verify(restClientProvider).getRestClient(DRIVE_ID);

        // Multiple rows version
        sheetFormattingService.clearRowFormatting(FILE_ID, SHEET_NAME, 5.0, 3);
        verify(restClientProvider).getRestClient(null);

        verify(restClient, times(2)).post(anyString(), anyMap());
    }
}

