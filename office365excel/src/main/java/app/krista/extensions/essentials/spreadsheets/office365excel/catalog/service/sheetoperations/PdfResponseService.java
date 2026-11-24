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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class PdfResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(PdfResponseService.class);

    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    private final DriveService driveService;

    @Inject
    public PdfResponseService(ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager,
                              ExtensionResponseGenerator responseGenerator,RequestContext requestContext,
                              TelemetryHelper telemetryHelper, DriveService driveService) {
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
        this.driveService = driveService;
    }

    public ExtensionResponse getUploadFileAndConvertToPDFExtensionResponse(File excelFile, String filePath, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.uploadFileAndConvertToPDF");

        AreaOperationHandler<Map<String, Object>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);

        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createUploadFileAndConvertToPDFValidationMap(excelFile);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createUploadFileAndConvertToPDFOperationTags(excelFile, filePath);

        return handler.execute(
                "uploadFileAndConvertToPDF",
                validationMap,
                Map.of(EXCEL_FILE, excelFile, FILE_PATH, filePath != null ? filePath : ""),
                SubCatalogConstants.CONFIRM_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF,
                stateId -> StateMapperUtil.uploadFileAndConvertToPDFMetaToMap(excelFile, filePath, stateId),
                () -> this.driveService.getPdfFileForGivenExcel(excelFile, filePath),
                result -> {
                    LOGGER.info("Successfully uploaded Excel file and converted to PDF");
                    return result;
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getConvertToPDFExtensionResponse(String fileId, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.convertToPDF");

        AreaOperationHandler<File> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createConvertToPdfValidationMapForSheet(fileId);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createConvertToPdfOperationTags(null, fileId);

        return handler.execute(
                "convertToPDF",
                validationMap,
                Map.of(FILE_ID, fileId),
                SubCatalogConstants.CONFIRM_REENTER_CONVERT_TO_PDF,
                stateId -> StateMapperUtil.convertToPdfMetaToMap(null, fileId, stateId),
                () -> {
                    try {
                        return this.driveService.getPdfFile(fileId, null);
                    } catch (IOException cause) {
                        throw new RuntimeException(cause);
                    }
                },
                result -> {
                    LOGGER.info("Successfully converted Excel file to PDF");
                    return Map.of("PDF File", result);
                },
                telemetryTags,
                allowRetry
        );
    }

}
