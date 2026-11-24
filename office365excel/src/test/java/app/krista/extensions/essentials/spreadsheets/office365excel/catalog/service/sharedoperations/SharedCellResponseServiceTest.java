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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.FreeForm;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("SharedCellResponseService Tests")
class SharedCellResponseServiceTest {

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

    private SharedCellResponseService service;

    private static final String DRIVE_ID = "driveId";
    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final String CELL_ID = "A1";
    private static final String DATA = "testData";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new SharedCellResponseService(worksheetService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getAddDataToCellExtensionResponse returns a non-null response")
    void testGetAddDataToCellExtensionResponse() {
        ExtensionResponse result = service.getAddDataToCellExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, DATA, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addDataToProvidedCell");
    }

    @Test
    @DisplayName("getCellDataExtensionResponse returns a non-null response")
    void testGetCellDataExtensionResponse() {
        ExtensionResponse result = service.getCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.cellData");
    }

    @Test
    @DisplayName("getAddMultipleCellDataExtensionResponse returns a non-null response")
    void testGetAddMultipleCellDataExtensionResponse() {
        List<Map<String, Object>> cellIDData = List.of(Map.of("cellId", "A1", "data", "val1"));

        ExtensionResponse result = service.getAddMultipleCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addMultipleCellData");
    }

    @Test
    @DisplayName("getAddDataToCellExtensionResponse increments telemetry counter")
    void testAddDataToCellIncrementsTelemetry() {
        service.getAddDataToCellExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, DATA, ALLOW_RETRY);

        verify(telemetryHelper, times(1)).incrementCount("office365Excel.addDataToProvidedCell");
    }

    @Test
    @DisplayName("getCellDataExtensionResponse increments telemetry counter")
    void testCellDataIncrementsTelemetry() {
        service.getCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, ALLOW_RETRY);

        verify(telemetryHelper, times(1)).incrementCount("office365Excel.cellData");
    }

    @Test
    @DisplayName("getAddMultipleCellDataExtensionResponse increments telemetry counter")
    void testAddMultipleCellDataIncrementsTelemetry() {
        List<Map<String, Object>> cellIDData = List.of(Map.of("cellId", "A1", "data", "val1"));

        service.getAddMultipleCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);

        verify(telemetryHelper, times(1)).incrementCount("office365Excel.addMultipleCellData");
    }

    @Test
    @DisplayName("getAddDataToCellExtensionResponse with allowRetry false")
    void testGetAddDataToCellExtensionResponseNoRetry() {
        ExtensionResponse result = service.getAddDataToCellExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, DATA, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addDataToProvidedCell");
    }

    @Test
    @DisplayName("getCellDataExtensionResponse with allowRetry false")
    void testGetCellDataExtensionResponseNoRetry() {
        ExtensionResponse result = service.getCellDataExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, CELL_ID, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.cellData");
    }
}
