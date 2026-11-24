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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.ExtensionResponseFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.RemediationActionFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.SharedFileReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.DriveService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.FileNotFoundException;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

/**
 * Implementation for shared file-specific reentry operations
 */
public class SharedFileReentryHandlerImpl extends BaseReentryHandlerImpl implements SharedFileReentryHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(SharedFileReentryHandlerImpl.class);

    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final DriveService driveService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public SharedFileReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                        DriveService driveService, TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.driveService = driveService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetSharedExcelFileConfirmation(Map<String, Object> map) {
        LOGGER.info("SharedFileReentryHandler.handleGetSharedExcelFileConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter shared Excel file data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("Reenter was true, hence continuing");
        telemetryHelper.incrementCount("office365Excel.getSharedExcelFileConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_SHARED_EXCEL_FILE,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetSharedExcelFileReentry(Map<String, Object> map) {
        LOGGER.info("SharedFileReentryHandler.handleGetSharedExcelFileReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetSharedExcelFileOperationTags(
                (String) map.get(SHARED_FOLDER_NAME), (String) map.get(SHARED_FILE_NAME));
        telemetryHelper.incrementCount("office365Excel.getSharedExcelFileReentry");

        try {
            return processSharedExcelFileReentry(map);
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getSharedExcelFileReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get shared Excel file on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving the shared Excel file. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getSharedExcelFileReentry", startTime, telemetryTags);
        }
    }

    private ExtensionResponse processSharedExcelFileReentry(Map<String, Object> map) {
        String sharedFolderName = (String) map.get(SHARED_FOLDER_NAME);
        String sharedFileName = (String) map.get(SHARED_FILE_NAME);

        LOGGER.info("Processing corrected input for get shared Excel file operation");

        validateExcelFileName(sharedFileName);
        List<ExcelFile> excelFiles = retrieveSharedExcelFile(sharedFolderName, sharedFileName);
        return createSuccessResponse(excelFiles, sharedFileName);
    }

    private void validateExcelFileName(String sharedFileName) {
        if (!StringUtil.isExcelFile(sharedFileName)) {
            LOGGER.debug("Invalid Excel file name on retry: {}", sharedFileName);
            throw new IllegalArgumentException("Not a valid Excel file: " + sharedFileName);
        }
    }

    private List<ExcelFile> retrieveSharedExcelFile(String sharedFolderName, String sharedFileName) {
        List<ExcelFile> excelFileFromSharedDrive = this.driveService.getExcelFileFromSharedDrive(
                sharedFolderName, sharedFileName);

        if (excelFileFromSharedDrive == null || excelFileFromSharedDrive.isEmpty()) {
            LOGGER.debug("No Excel file found with name: {}", sharedFileName);
            throw new RuntimeException(new FileNotFoundException("Excel file not found: " + sharedFileName));
        }

        return excelFileFromSharedDrive;
    }

    private ExtensionResponse createSuccessResponse(List<ExcelFile> excelFiles, String sharedFileName) {
        LOGGER.info("Successfully retrieved shared Excel file: {}", sharedFileName);
        return ExtensionResponseFactory.create(Map.of("Shared Excel File", excelFiles.getFirst()));
    }
}
