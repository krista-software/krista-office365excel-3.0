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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.MinMaxValueReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;

/**
 * Implementation for min-max value reentry operations
 */
public class MinMaxValueReentryHandlerImpl extends BaseReentryHandlerImpl implements MinMaxValueReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(MinMaxValueReentryHandlerImpl.class);

    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final WorksheetService worksheetService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public MinMaxValueReentryHandlerImpl(ErrorHandlingStateManager internalStateManager,
                                         ExtensionResponseGenerator responseGenerator,
                                         WorksheetService worksheetService,
                                         TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.worksheetService = worksheetService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleConfirmation(Map<String, Object> map) {
        LOGGER.info("MinMaxValueReentryHandler.handleConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter min-max value data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter min-max value data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.minMaxValueConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_GET_MIN_MAX_VALUE,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReentry(Map<String, Object> map) {
        LOGGER.info("MinMaxValueReentryHandler.handleReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createMinMaxValueOperationTags((String) map.get(KEY_COLUMN_NAME),
                (String) map.get(COLUMN_RANGE));
        telemetryHelper.incrementCount("office365Excel.minMaxValueReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String keyColumnName = (String) map.get(KEY_COLUMN_NAME);
            String columnRange = (String) map.get(COLUMN_RANGE);

            LOGGER.info("Processing corrected input for min-max value operation");
            Map<String, Object> result = this.worksheetService.findMinMaxValueFromRowDataByColumnRange(
                    null, fileId, sheetName, keyColumnName, columnRange);

            LOGGER.info("Successfully found min-max values for column range: {}", columnRange);
            return ExtensionResponseFactory.create(Map.of("Result", result));
        } catch (Exception cause) {
            LOGGER.error("Error in min-max value reentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.minMaxValueReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get min-max values on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction(
                            "An error occurred while retrieving min-max values. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.minMaxValueReentry", startTime, telemetryTags);
        }
    }
}