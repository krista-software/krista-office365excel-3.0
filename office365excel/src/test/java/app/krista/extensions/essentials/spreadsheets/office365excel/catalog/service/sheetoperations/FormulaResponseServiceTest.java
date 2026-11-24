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
import app.krista.ksdk.context.RequestContext;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("FormulaResponseService Tests")
class FormulaResponseServiceTest {

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

    private FormulaResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new FormulaResponseService(validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper, workbookService);
    }

    @Test
    @DisplayName("getExecuteFormulaExtensionResponse returns non-null response and increments telemetry")
    void testGetExecuteFormulaExtensionResponse() {
        String formula = "=SUM(A1:A10)";
        String formulaName = "SUM";
        String comment = "test";

        ExtensionResponse result = service.getExecuteFormulaExtensionResponse(FILE_ID, formula, formulaName, comment, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.executeFormula");
    }

    @Test
    @DisplayName("getVlookupExtensionResponse returns non-null response and increments telemetry")
    void testGetVlookupExtensionResponse() {
        String lookupValue = "test";
        String range = "A1:B10";
        Double columnIndex = 2.0;
        Boolean exactMatch = true;

        ExtensionResponse result = service.getVlookupExtensionResponse(FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.vlookup");
    }

    @Test
    @DisplayName("getSummationExtensionResponse returns non-null response and increments telemetry")
    void testGetSummationExtensionResponse() {
        String range = "A1:A10";

        ExtensionResponse result = service.getSummationExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.summation");
    }

    @Test
    @DisplayName("getExecuteFormulaExtensionResponse with allowRetry false")
    void testGetExecuteFormulaExtensionResponseNoRetry() {
        String formula = "=AVERAGE(B1:B10)";
        String formulaName = "AVERAGE";
        String comment = "calc";

        ExtensionResponse result = service.getExecuteFormulaExtensionResponse(FILE_ID, formula, formulaName, comment, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.executeFormula");
    }

    @Test
    @DisplayName("getVlookupExtensionResponse with null exactMatch")
    void testGetVlookupExtensionResponseNullExactMatch() {
        String lookupValue = "value";
        String range = "A1:C20";
        Double columnIndex = 3.0;

        ExtensionResponse result = service.getVlookupExtensionResponse(FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.vlookup");
    }

    @Test
    @DisplayName("getVlookupExtensionResponse with exactMatch false")
    void testGetVlookupExtensionResponseExactMatchFalse() {
        String lookupValue = "approxValue";
        String range = "A1:D10";
        Double columnIndex = 4.0;
        Boolean exactMatch = false;

        ExtensionResponse result = service.getVlookupExtensionResponse(FILE_ID, lookupValue, SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.vlookup");
    }

    @Test
    @DisplayName("getSummationExtensionResponse with different range")
    void testGetSummationExtensionResponseDifferentRange() {
        String range = "B5:B50";

        ExtensionResponse result = service.getSummationExtensionResponse(FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.summation");
    }
}
