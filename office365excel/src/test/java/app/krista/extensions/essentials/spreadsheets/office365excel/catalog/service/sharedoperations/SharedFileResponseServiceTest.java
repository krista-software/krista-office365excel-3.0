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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.DriveService;
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
@DisplayName("SharedFileResponseService Tests")
class SharedFileResponseServiceTest {

    @Mock
    private WorkbookService workbookService;

    @Mock
    private DriveService driveService;

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

    private SharedFileResponseService service;

    private static final String DRIVE_ID = "driveId";
    private static final String FILE_ID = "fileId";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new SharedFileResponseService(workbookService, driveService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getSharedExcelFileExtensionResponse returns non-null response")
    void testGetSharedExcelFileExtensionResponse() {
        String folderName = "SharedFolder";
        String fileName = "test.xlsx";

        ExtensionResponse result = service.getSharedExcelFileExtensionResponse(folderName, fileName, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getSharedExcelFile");
    }

    @Test
    @DisplayName("getSharedExcelFileExtensionResponse with null folder name")
    void testGetSharedExcelFileExtensionResponseNullFolder() {
        String fileName = "test.xlsx";

        ExtensionResponse result = service.getSharedExcelFileExtensionResponse(null, fileName, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getSharedExcelFile");
    }

    @Test
    @DisplayName("getConvertToPdfExtensionResponse returns non-null response")
    void testGetConvertToPdfExtensionResponse() {
        ExtensionResponse result = service.getConvertToPdfExtensionResponse(DRIVE_ID, FILE_ID, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.convertToPDF");
    }

    @Test
    @DisplayName("getConvertToPdfExtensionResponse with allowRetry false")
    void testGetConvertToPdfExtensionResponseNoRetry() {
        ExtensionResponse result = service.getConvertToPdfExtensionResponse(DRIVE_ID, FILE_ID, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.convertToPDF");
    }

    @Test
    @DisplayName("getExecuteFormulaExtensionResponse returns non-null response")
    void testGetExecuteFormulaExtensionResponse() {
        String formula = "=SUM(A1:A10)";
        String formulaName = "SUM";
        String comment = "test";

        ExtensionResponse result = service.getExecuteFormulaExtensionResponse(DRIVE_ID, FILE_ID, formula, formulaName, comment, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.executeFormulaForSharedFile");
    }

    @Test
    @DisplayName("getExecuteFormulaExtensionResponse increments correct telemetry counter")
    void testGetExecuteFormulaIncrementsTelemetry() {
        String formula = "=AVERAGE(B1:B10)";
        String formulaName = "AVERAGE";
        String comment = "calc";

        service.getExecuteFormulaExtensionResponse(DRIVE_ID, FILE_ID, formula, formulaName, comment, ALLOW_RETRY);

        verify(telemetryHelper, times(1)).incrementCount("office365Excel.executeFormulaForSharedFile");
    }

    @Test
    @DisplayName("getSharedExcelFileExtensionResponse increments correct telemetry counter")
    void testGetSharedExcelFileIncrementsTelemetry() {
        service.getSharedExcelFileExtensionResponse("folder", "file.xlsx", ALLOW_RETRY);

        verify(telemetryHelper, times(1)).incrementCount("office365Excel.getSharedExcelFile");
    }
}
