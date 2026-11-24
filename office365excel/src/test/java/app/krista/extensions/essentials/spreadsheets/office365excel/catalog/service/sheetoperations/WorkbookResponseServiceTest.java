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
@DisplayName("WorkbookResponseService Tests")
class WorkbookResponseServiceTest {

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

    private WorkbookResponseService service;

    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new WorkbookResponseService(validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper, workbookService);
    }

    @Test
    @DisplayName("getGetExcelFileInformationExtensionResponse returns non-null response and increments telemetry")
    void testGetGetExcelFileInformationExtensionResponse() {
        String path = "/documents";
        String name = "test.xlsx";

        ExtensionResponse result = service.getGetExcelFileInformationExtensionResponse(path, name, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getExcelFileInformation");
    }

    @Test
    @DisplayName("getGetExcelFileInformationExtensionResponse with allowRetry false")
    void testGetGetExcelFileInformationExtensionResponseNoRetry() {
        String path = "/reports";
        String name = "report.xlsx";

        ExtensionResponse result = service.getGetExcelFileInformationExtensionResponse(path, name, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getExcelFileInformation");
    }

    @Test
    @DisplayName("getGetExcelFileInformationExtensionResponse with root path")
    void testGetGetExcelFileInformationExtensionResponseRootPath() {
        String path = "/";
        String name = "workbook.xlsx";

        ExtensionResponse result = service.getGetExcelFileInformationExtensionResponse(path, name, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getExcelFileInformation");
    }

    @Test
    @DisplayName("getGetExcelFileInformationExtensionResponse with nested path")
    void testGetGetExcelFileInformationExtensionResponseNestedPath() {
        String path = "/documents/finance/reports";
        String name = "quarterly.xlsx";

        ExtensionResponse result = service.getGetExcelFileInformationExtensionResponse(path, name, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getExcelFileInformation");
    }
}
