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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.WorkbookReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

/**
 * Implementation for workbook-specific reentry operations
 */
public class WorkbookReentryHandlerImpl extends BaseReentryHandlerImpl implements WorkbookReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookReentryHandlerImpl.class);
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final WorkbookService workbookService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public WorkbookReentryHandlerImpl(ErrorHandlingStateManager internalStateManager,ExtensionResponseGenerator responseGenerator,
                                      WorkbookService workbookService, TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.workbookService = workbookService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetExcelFileInfoConfirmation(Map<String, Object> inputMap) {
        LOGGER.info("WorkbookReentryHandler.handleGetExcelFileInfoConfirmation start: {}", inputMap);
        Boolean reenter = (Boolean) inputMap.get(REENTER);
        String stateId = (String) inputMap.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter Excel file information. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter Excel file information. Returning fetch prompt response.");
        telemetryHelper.incrementCount("office365Excel.getExcelFileInfoConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_GET_EXCEL_FILE_INFO,
                inputMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetExcelFileInfoReentry(Map<String, Object> inputMap) {
        LOGGER.info("WorkbookReentryHandler.handleGetExcelFileInfoReentry start: {}", inputMap);
        String stateId = (String) inputMap.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetExcelFileInfoOperationTags(
                (String) inputMap.get(PATH), (String) inputMap.get(NAME));
        telemetryHelper.incrementCount("office365Excel.getExcelFileInfoReentry");

        try {
            String path = (String) inputMap.get(PATH);
            String name = (String) inputMap.get(NAME);

            LOGGER.info("Processing corrected input for path: {}, name: {}", path, name);
            ExcelFile result = this.workbookService.getWorkbookDetails(path, name);

            LOGGER.info("Successfully retrieved Excel file information for: {}", name);
            return ExtensionResponseFactory.create(Map.of("Excel File", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getExcelFileInfoReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get Excel file information on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving the Excel file information. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getExcelFileInfoReentry", startTime, telemetryTags);
        }
    }
}