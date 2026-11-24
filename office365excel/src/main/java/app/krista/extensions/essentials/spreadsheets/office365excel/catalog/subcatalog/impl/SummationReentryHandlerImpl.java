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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.SummationReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
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
 * Implementation for summation reentry operations
 */
public class SummationReentryHandlerImpl extends BaseReentryHandlerImpl implements SummationReentryHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(SummationReentryHandlerImpl.class);

    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final WorkbookService workbookService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public SummationReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                       WorkbookService workbookService, TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.workbookService = workbookService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleSummationConfirmation(Map<String, Object> map) {
        LOGGER.info("SummationReentryHandler.handleSummationConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter summation data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter summation data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.summationConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_SUMMATION,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleSummationReentry(Map<String, Object> map) {
        LOGGER.info("SummationReentryHandler.handleSummationReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags((String) map.get(SHEET_NAME), (String) map.get(RANGE));
        telemetryHelper.incrementCount("office365Excel.summationReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String range = (String) map.get(RANGE);

            LOGGER.info("Processing corrected input for summation operation");
            Double result = this.workbookService.worksheetRangeSum(null, fileId, sheetName, range);

            LOGGER.info("Successfully calculated summation for range: {}", range);
            return ExtensionResponseFactory.create(Map.of("Sum", result));
        } catch (Exception cause) {
            LOGGER.error("Error in summation reentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.summationReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to calculate sum on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while calculating the sum. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.summationReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleTableSummationConfirmation(Map<String, Object> map) {
        LOGGER.info("TableSummationReentryHandler.handleTableSummationConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter table summation data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter table summation data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.tableSummationConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_TABLE_SUMMATION,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleTableSummationReentry(Map<String, Object> map) {
        LOGGER.info("TableSummationReentryHandler.handleTableSummationReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        List<String> columns = (List<String>) map.get(COLUMNS);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableSummationOperationTags(
                (String) map.get(TABLE_NAME), columns);
        telemetryHelper.incrementCount("office365Excel.tableSummationReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);

            LOGGER.info("Processing corrected input for table summation operation");
            Double result = this.workbookService.tableColumnsSum(null, fileId, tableName, columns);

            LOGGER.info("Successfully calculated table summation for table: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Sum", result));
        } catch (Exception cause) {
            LOGGER.error("Error in table summation reentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.tableSummationReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to calculate table sum on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while calculating the sum. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.tableSummationReentry", startTime, telemetryTags);
        }
    }
}