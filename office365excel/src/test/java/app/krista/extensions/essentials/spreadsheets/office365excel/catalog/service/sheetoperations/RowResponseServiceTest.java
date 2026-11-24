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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.SheetFormattingService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
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
@DisplayName("RowResponseService Tests")
class RowResponseServiceTest {

    @Mock
    private WorksheetService worksheetService;

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
    private StringUtil stringUtil;

    @Mock
    private SheetFormattingService sheetFormattingService;

    private RowResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new RowResponseService(worksheetService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper,
                restClientProvider, stringUtil, sheetFormattingService);
    }

    @Test
    @DisplayName("getAddRowsExtensionResponse returns non-null response and increments telemetry")
    void testGetAddRowsExtensionResponse() {
        List<String> rows = List.of("row1", "row2");
        List<File> rowsFromCsv = null;

        ExtensionResponse result = service.getAddRowsExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRows");
    }

    @Test
    @DisplayName("getAddRowsExtensionResponse with null rows")
    void testGetAddRowsExtensionResponseNullRows() {
        ExtensionResponse result = service.getAddRowsExtensionResponse(FILE_ID, SHEET_NAME, null, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRows");
    }

    @Test
    @DisplayName("getAddRowWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetAddRowWithIndexExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsv = null;
        Double rowIndex = 3.0;

        ExtensionResponse result = service.getAddRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowWithIndex");
    }

    @Test
    @DisplayName("getAddRowWithIndexExtensionResponse with null rows")
    void testGetAddRowWithIndexExtensionResponseNullRows() {
        Double rowIndex = 1.0;

        ExtensionResponse result = service.getAddRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, null, null, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowWithIndex");
    }

    @Test
    @DisplayName("getGetCompleteRowInfoExtensionResponse returns non-null response and increments telemetry")
    void testGetGetCompleteRowInfoExtensionResponse() {
        Double rowIndex = 5.0;

        ExtensionResponse result = service.getGetCompleteRowInfoExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getCompleteRowInfo");
    }

    @Test
    @DisplayName("getGetRowsForRangeExtensionResponse returns non-null response and increments telemetry")
    void testGetGetRowsForRangeExtensionResponse() {
        String range = "A1:C5";

        ExtensionResponse result = service.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getRowsForRange");
    }

    @Test
    @DisplayName("getDeleteRowWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetDeleteRowWithIndexExtensionResponse() {
        Double rowIndex = 4.0;

        ExtensionResponse result = service.getDeleteRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteRowWithIndex");
    }

    @Test
    @DisplayName("getUpdateRowWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetUpdateRowWithIndexExtensionResponse() {
        String row = "updatedRow";
        Double rowIndex = 2.0;

        ExtensionResponse result = service.getUpdateRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.updateRowWithIndex");
    }

    @Test
    @DisplayName("getDeleteRowWithIndexExtensionResponse with allowRetry false")
    void testGetDeleteRowWithIndexExtensionResponseNoRetry() {
        Double rowIndex = 1.0;

        ExtensionResponse result = service.getDeleteRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteRowWithIndex");
    }

    @Test
    @DisplayName("getGetCompleteRowInfoExtensionResponse with allowRetry false")
    void testGetGetCompleteRowInfoExtensionResponseNoRetry() {
        Double rowIndex = 10.0;

        ExtensionResponse result = service.getGetCompleteRowInfoExtensionResponse(FILE_ID, SHEET_NAME, rowIndex, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getCompleteRowInfo");
    }

    @Test
    @DisplayName("getUpdateRowWithIndexExtensionResponse with allowRetry false")
    void testGetUpdateRowWithIndexExtensionResponseNoRetry() {
        String row = "data1,data2";
        Double rowIndex = 7.0;

        ExtensionResponse result = service.getUpdateRowWithIndexExtensionResponse(FILE_ID, SHEET_NAME, row, rowIndex, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.updateRowWithIndex");
    }
}
