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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.FindValueReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;

/**
 * Implementation for find value reentry operations
 */
public class FindValueReentryHandlerImpl extends BaseReentryHandlerImpl implements FindValueReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(FindValueReentryHandlerImpl.class);

    private final WorksheetService worksheetService;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final TelemetryHelper telemetryHelper;

    /**
     * Constructor for FindValueReentryHandlerImpl
     *
     * @param internalStateManager Internal state manager
     * @param responseGenerator    Response generator
     * @param worksheetService     Worksheet service
     * @param telemetryHelper      Telemetry helper
     */
    @Inject
    public FindValueReentryHandlerImpl(ErrorHandlingStateManager internalStateManager,
                                       ExtensionResponseGenerator responseGenerator,
                                       WorksheetService worksheetService,
                                       TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.worksheetService = worksheetService;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleConfirmation(Map<String, Object> map) {
        LOGGER.info("FindValueReentryHandler.handleConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter find value data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter find value data. Returning reentry form.");
        telemetryHelper.incrementCount("office365Excel.findValueConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_FIND_VALUE,
                Map.of(STATE_ID, stateId));
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReentry(Map<String, Object> map) {
        LOGGER.info("FindValueReentryHandler.handleReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createFindValueOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(LOOKUP_VALUE));
        telemetryHelper.incrementCount("office365Excel.findValueReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String lookupValue = (String) map.get(LOOKUP_VALUE);

            LOGGER.info("Processing corrected input for find value operation");
            List<String> result = this.worksheetService.findValue(null, fileId, sheetName, lookupValue);

            LOGGER.info("Successfully found values for lookup: {}", lookupValue);
            return ExtensionResponseFactory.create(Map.of("Found Values", result));
        } catch (Exception cause) {
            LOGGER.error("Error in find value reentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.findValueReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to find value on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    null, null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.findValueReentry", startTime, telemetryTags);
        }
    }
}