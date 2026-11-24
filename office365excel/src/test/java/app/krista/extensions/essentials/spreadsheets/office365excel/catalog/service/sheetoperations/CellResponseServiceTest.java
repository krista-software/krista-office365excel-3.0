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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
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
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("CellResponseService Tests")
class CellResponseServiceTest {

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
    private WorkbookService workbookService;

    private CellResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final String CELL_ID = "A1";
    private static final String DATA = "testData";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new CellResponseService(worksheetService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper, workbookService);
    }

    @Test
    @DisplayName("getGetCellValueExtensionResponse returns non-null response and increments telemetry")
    void testGetGetCellValueExtensionResponse() {
        ExtensionResponse result = service.getGetCellValueExtensionResponse(FILE_ID, SHEET_NAME, CELL_ID, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getCellValue");
    }

    @Test
    @DisplayName("getAddDataToCellExtensionResponse returns non-null response and increments telemetry")
    void testGetAddDataToCellExtensionResponse() {
        ExtensionResponse result = service.getAddDataToCellExtensionResponse(FILE_ID, SHEET_NAME, CELL_ID, DATA, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addDataToCell");
    }

    @Test
    @DisplayName("getAddMultipleCellDataExtensionResponse returns non-null response and increments telemetry")
    void testGetAddMultipleCellDataExtensionResponse() {
        List<Map<String, Object>> cellIDData = List.of(Map.of("cellId", "A1", "data", "val1"));

        ExtensionResponse result = service.getAddMultipleCellDataExtensionResponse(FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addMultipleCellData");
    }

    @Test
    @DisplayName("getAddFormulaToCellExtensionResponse returns non-null response and increments telemetry")
    void testGetAddFormulaToCellExtensionResponse() {
        String formula = "=SUM(A1:A10)";

        ExtensionResponse result = service.getAddFormulaToCellExtensionResponse(FILE_ID, SHEET_NAME, CELL_ID, formula, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addFormulaToCell");
    }

    @Test
    @DisplayName("getGetCellValueExtensionResponse with allowRetry false")
    void testGetGetCellValueExtensionResponseNoRetry() {
        ExtensionResponse result = service.getGetCellValueExtensionResponse(FILE_ID, SHEET_NAME, CELL_ID, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getCellValue");
    }

    @Test
    @DisplayName("getAddDataToCellExtensionResponse with allowRetry false")
    void testGetAddDataToCellExtensionResponseNoRetry() {
        ExtensionResponse result = service.getAddDataToCellExtensionResponse(FILE_ID, SHEET_NAME, CELL_ID, DATA, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addDataToCell");
    }

    @Test
    @DisplayName("getAddFormulaToCellExtensionResponse with different formula")
    void testGetAddFormulaToCellExtensionResponseDifferentFormula() {
        String formula = "=VLOOKUP(A1,B1:D10,3,FALSE)";

        ExtensionResponse result = service.getAddFormulaToCellExtensionResponse(FILE_ID, SHEET_NAME, "B5", formula, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addFormulaToCell");
    }

    @Test
    @DisplayName("getAddMultipleCellDataExtensionResponse with multiple cells")
    void testGetAddMultipleCellDataExtensionResponseMultipleCells() {
        List<Map<String, Object>> cellIDData = List.of(
                Map.of("cellId", "A1", "data", "val1"),
                Map.of("cellId", "B2", "data", "val2"),
                Map.of("cellId", "C3", "data", "val3")
        );

        ExtensionResponse result = service.getAddMultipleCellDataExtensionResponse(FILE_ID, SHEET_NAME, cellIDData, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addMultipleCellData");
    }
}
