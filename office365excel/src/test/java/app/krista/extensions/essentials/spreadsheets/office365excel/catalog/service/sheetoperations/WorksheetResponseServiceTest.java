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
@DisplayName("WorksheetResponseService Tests")
class WorksheetResponseServiceTest {

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

    private WorksheetResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new WorksheetResponseService(worksheetService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getDeleteSheetExtensionResponse returns non-null response and increments telemetry")
    void testGetDeleteSheetExtensionResponse() {
        ExtensionResponse result = service.getDeleteSheetExtensionResponse(FILE_ID, SHEET_NAME, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteSheet");
    }

    @Test
    @DisplayName("getCreateSheetExtensionResponse returns non-null response and increments telemetry")
    void testGetCreateSheetExtensionResponse() {
        String newSheetName = "NewSheet";

        ExtensionResponse result = service.getCreateSheetExtensionResponse(FILE_ID, newSheetName, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.createSheet");
    }

    @Test
    @DisplayName("getGetSheetsInformationExtensionResponse returns non-null response and increments telemetry")
    void testGetGetSheetsInformationExtensionResponse() {
        ExtensionResponse result = service.getGetSheetsInformationExtensionResponse(FILE_ID, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getSheetsInformation");
    }

    @Test
    @DisplayName("getClearSheetForRangeExtensionResponse returns non-null response and increments telemetry")
    void testGetClearSheetForRangeExtensionResponse() {
        String cellRange = "A1:Z100";

        ExtensionResponse result = service.getClearSheetForRangeExtensionResponse(FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.clearSheetForRange");
    }

    @Test
    @DisplayName("getDeleteSheetExtensionResponse with allowRetry false")
    void testGetDeleteSheetExtensionResponseNoRetry() {
        ExtensionResponse result = service.getDeleteSheetExtensionResponse(FILE_ID, SHEET_NAME, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteSheet");
    }

    @Test
    @DisplayName("getCreateSheetExtensionResponse with allowRetry false")
    void testGetCreateSheetExtensionResponseNoRetry() {
        String newSheetName = "TestSheet";

        ExtensionResponse result = service.getCreateSheetExtensionResponse(FILE_ID, newSheetName, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.createSheet");
    }

    @Test
    @DisplayName("getGetSheetsInformationExtensionResponse with allowRetry false")
    void testGetGetSheetsInformationExtensionResponseNoRetry() {
        ExtensionResponse result = service.getGetSheetsInformationExtensionResponse(FILE_ID, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getSheetsInformation");
    }

    @Test
    @DisplayName("getClearSheetForRangeExtensionResponse with specific range")
    void testGetClearSheetForRangeExtensionResponseSpecificRange() {
        String cellRange = "B2:D5";

        ExtensionResponse result = service.getClearSheetForRangeExtensionResponse(FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.clearSheetForRange");
    }
}
