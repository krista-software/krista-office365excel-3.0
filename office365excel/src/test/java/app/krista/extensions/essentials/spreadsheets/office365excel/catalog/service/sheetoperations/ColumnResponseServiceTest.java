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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sheetoperations;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
import app.krista.ksdk.context.RequestContext;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("ColumnResponseService Tests")
class ColumnResponseServiceTest {

    @Mock
    private ValidationOrchestrator validationOrchestrator;

    @Mock
    private ErrorHandlingStateManager internalStateManager;

    @Mock
    private ExtensionResponseGenerator responseGenerator;

    @Mock
    private RequestContext requestContext;

    @Mock
    private TelemetryHelper telemetryHelper;

    @Mock
    private RestClientProvider restClientProvider;

    @Mock
    private WorksheetOperation worksheetOperation;

    private ColumnResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new ColumnResponseService(validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper, restClientProvider, worksheetOperation);
    }

    @Test
    @DisplayName("getDeleteColumnWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetDeleteColumnWithIndexExtensionResponse() {
        String columnName = "B";

        ExtensionResponse result = service.getDeleteColumnWithIndexExtensionResponse(FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteColumnWithIndex");
    }

    @Test
    @DisplayName("getGetColumnValuesForKeysInSheetExtensionResponse returns non-null response and increments telemetry")
    void testGetGetColumnValuesForKeysInSheetExtensionResponse() {
        String range = "A1:D10";
        List<String> keyRows = List.of("key1");
        List<String> keyColumns = List.of("col1");
        Boolean getAllRecords = false;
        Double columnNumber = 1.0;

        ExtensionResponse result = service.getGetColumnValuesForKeysInSheetExtensionResponse(FILE_ID, SHEET_NAME, range, keyRows, keyColumns, getAllRecords, columnNumber, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeysInSheet");
    }

    @Test
    @DisplayName("getDeleteColumnWithIndexExtensionResponse with allowRetry false")
    void testGetDeleteColumnWithIndexExtensionResponseNoRetry() {
        String columnName = "C";

        ExtensionResponse result = service.getDeleteColumnWithIndexExtensionResponse(FILE_ID, SHEET_NAME, columnName, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteColumnWithIndex");
    }

    @Test
    @DisplayName("getGetColumnValuesForKeysInSheetExtensionResponse with null columnNumber")
    void testGetGetColumnValuesForKeysInSheetExtensionResponseNullColumnNumber() {
        String range = "A1:D10";
        List<String> keyRows = List.of("key1");
        List<String> keyColumns = List.of("col1");
        Boolean getAllRecords = true;

        ExtensionResponse result = service.getGetColumnValuesForKeysInSheetExtensionResponse(FILE_ID, SHEET_NAME, range, keyRows, keyColumns, getAllRecords, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeysInSheet");
    }

    @Test
    @DisplayName("getGetColumnValuesForKeysInSheetExtensionResponse with getAllRecords true")
    void testGetGetColumnValuesForKeysInSheetExtensionResponseGetAllRecords() {
        String range = "A1:Z50";
        List<String> keyRows = List.of("row1", "row2");
        List<String> keyColumns = List.of("col1", "col2");
        Boolean getAllRecords = true;
        Double columnNumber = 3.0;

        ExtensionResponse result = service.getGetColumnValuesForKeysInSheetExtensionResponse(FILE_ID, SHEET_NAME, range, keyRows, keyColumns, getAllRecords, columnNumber, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeysInSheet");
    }
}
