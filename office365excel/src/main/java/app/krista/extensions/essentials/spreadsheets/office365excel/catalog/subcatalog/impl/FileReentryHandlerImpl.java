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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.impl;

import app.krista.extension.executor.ExtensionResponse;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.ExtensionResponseFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.RemediationActionFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.FileReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.DriveService;
import app.krista.model.base.File;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

/**
 * Implementation for file-specific reentry operations
 */
public class FileReentryHandlerImpl extends BaseReentryHandlerImpl implements FileReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(FileReentryHandlerImpl.class);

    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final DriveService driveService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public FileReentryHandlerImpl(ErrorHandlingStateManager internalStateManager,
                                  ExtensionResponseGenerator responseGenerator,
                                  DriveService driveService,
                                  TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.driveService = driveService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleUploadFileAndConvertToPDFConfirmation(Map<String, Object> map) {
        LOGGER.info("FileReentryHandler.handleUploadFileAndConvertToPDFConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter upload file data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter upload file and convert to PDF data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.uploadFileAndConvertToPDFConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleUploadFileAndConvertToPDFReentry(Map<String, Object> map) {
        LOGGER.info("FileReentryHandler.handleUploadFileAndConvertToPDFReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createUploadFileAndConvertToPDFOperationTags(
                (String) map.get(FILE_PATH));
        telemetryHelper.incrementCount("office365Excel.uploadFileAndConvertToPDFReentry");

        try {
            File excelFile = (File) map.get(EXCEL_FILE);
            String filePath = (String) map.get(FILE_PATH);

            LOGGER.info("Processing corrected input for upload file and convert to PDF operation");
            Map<String, Object> result = this.driveService.getPdfFileForGivenExcel(excelFile, filePath);
            LOGGER.info("Successfully uploaded Excel file and converted to PDF");
            return ExtensionResponseFactory.create(Map.of("Data", result));
        } catch (Exception cause) {
            LOGGER.error("Error processing reentry for upload file and convert to PDF: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.uploadFileAndConvertToPDFReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to upload file and convert to PDF on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while uploading and converting file. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.uploadFileAndConvertToPDFReentry", startTime, telemetryTags);
        }
    }
}