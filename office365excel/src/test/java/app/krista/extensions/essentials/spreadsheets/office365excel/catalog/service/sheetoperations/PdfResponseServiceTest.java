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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.DriveService;
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

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("PdfResponseService Tests")
class PdfResponseServiceTest {

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
    private DriveService driveService;

    private PdfResponseService service;

    private static final String FILE_ID = "fileId";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new PdfResponseService(validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper, driveService);
    }

    @Test
    @DisplayName("getUploadFileAndConvertToPDFExtensionResponse returns non-null response and increments telemetry")
    void testGetUploadFileAndConvertToPDFExtensionResponse() {
        File excelFile = mock(File.class);
        String filePath = "/path/to/file.xlsx";

        ExtensionResponse result = service.getUploadFileAndConvertToPDFExtensionResponse(excelFile, filePath, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.uploadFileAndConvertToPDF");
    }

    @Test
    @DisplayName("getConvertToPDFExtensionResponse returns non-null response and increments telemetry")
    void testGetConvertToPDFExtensionResponse() {
        ExtensionResponse result = service.getConvertToPDFExtensionResponse(FILE_ID, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.convertToPDF");
    }

    @Test
    @DisplayName("getUploadFileAndConvertToPDFExtensionResponse with null filePath")
    void testGetUploadFileAndConvertToPDFExtensionResponseNullPath() {
        File excelFile = mock(File.class);

        ExtensionResponse result = service.getUploadFileAndConvertToPDFExtensionResponse(excelFile, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.uploadFileAndConvertToPDF");
    }

    @Test
    @DisplayName("getConvertToPDFExtensionResponse with allowRetry false")
    void testGetConvertToPDFExtensionResponseNoRetry() {
        ExtensionResponse result = service.getConvertToPDFExtensionResponse(FILE_ID, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.convertToPDF");
    }

    @Test
    @DisplayName("getUploadFileAndConvertToPDFExtensionResponse with allowRetry false")
    void testGetUploadFileAndConvertToPDFExtensionResponseNoRetry() {
        File excelFile = mock(File.class);
        String filePath = "/another/path/to/file.xlsx";

        ExtensionResponse result = service.getUploadFileAndConvertToPDFExtensionResponse(excelFile, filePath, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.uploadFileAndConvertToPDF");
    }
}
