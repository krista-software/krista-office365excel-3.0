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

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("SearchResponseService Tests")
class SearchResponseServiceTest {

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

    private SearchResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new SearchResponseService(worksheetService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getFindValueExtensionResponse returns non-null response and increments telemetry")
    void testGetFindValueExtensionResponse() {
        String lookupValue = "findMe";

        ExtensionResponse result = service.getFindValueExtensionResponse(FILE_ID, SHEET_NAME, lookupValue, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.findValue");
    }

    @Test
    @DisplayName("getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse returns non-null response and increments telemetry")
    void testGetGetMinMaxValueFromRowDataByColumnRangeExtensionResponse() {
        String keyColumnName = "Amount";
        String columnRange = "A1:A100";

        ExtensionResponse result = service.getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(FILE_ID, SHEET_NAME, keyColumnName, columnRange, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getMinMaxValueFromRowDataByColumnRange");
    }

    @Test
    @DisplayName("getFindValueExtensionResponse with allowRetry false")
    void testGetFindValueExtensionResponseNoRetry() {
        String lookupValue = "searchValue";

        ExtensionResponse result = service.getFindValueExtensionResponse(FILE_ID, SHEET_NAME, lookupValue, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.findValue");
    }

    @Test
    @DisplayName("getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse with allowRetry false")
    void testGetGetMinMaxValueFromRowDataByColumnRangeExtensionResponseNoRetry() {
        String keyColumnName = "Price";
        String columnRange = "B1:B50";

        ExtensionResponse result = service.getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(FILE_ID, SHEET_NAME, keyColumnName, columnRange, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getMinMaxValueFromRowDataByColumnRange");
    }

    @Test
    @DisplayName("getFindValueExtensionResponse with numeric lookup value")
    void testGetFindValueExtensionResponseNumericValue() {
        String lookupValue = "12345";

        ExtensionResponse result = service.getFindValueExtensionResponse(FILE_ID, SHEET_NAME, lookupValue, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.findValue");
    }

    @Test
    @DisplayName("getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse with wide column range")
    void testGetGetMinMaxValueFromRowDataByColumnRangeExtensionResponseWideRange() {
        String keyColumnName = "Revenue";
        String columnRange = "A1:Z1000";

        ExtensionResponse result = service.getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(FILE_ID, SHEET_NAME, keyColumnName, columnRange, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getMinMaxValueFromRowDataByColumnRange");
    }
}
