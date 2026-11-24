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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.ColumnValuesReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
import app.krista.model.base.FreeForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER;

/**
 * Implementation for column values reentry operations
 */
public class ColumnValuesReentryHandlerImpl extends BaseReentryHandlerImpl implements ColumnValuesReentryHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(ColumnValuesReentryHandlerImpl.class);

    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final WorksheetOperation worksheetOperation;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public ColumnValuesReentryHandlerImpl(ErrorHandlingStateManager internalStateManager,
                                          ExtensionResponseGenerator responseGenerator,
                                          WorksheetOperation worksheetOperation,
                                          TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.worksheetOperation = worksheetOperation;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetColumnValuesForKeysInSheetConfirmation(Map<String, Object> map) {
        LOGGER.info("ColumnValuesReentryHandler.handleGetColumnValuesForKeysInSheetConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter column values for keys in sheet data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter column values for keys in sheet data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeysInSheetConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetColumnValuesForKeysInSheetReentry(Map<String, Object> map) {
        LOGGER.info("ColumnValuesReentryHandler.handleGetColumnValuesForKeysInSheetReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetColumnValuesForKeysInSheetOperationTags(
                (String) map.get(SHEET_NAME), (List<String>) map.get(LOOKUP_ROWS), (List<String>) map.get(LOOKUP_COLUMNS));
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeysInSheetReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String range = (String) map.get(RANGE);
            List<String> lookupRows = (List<String>) map.get(LOOKUP_ROWS);
            List<String> lookupColumns = (List<String>) map.get(LOOKUP_COLUMNS);
            Boolean getAllRecord = (Boolean) map.get(GET_ALL_RECORD);
            Double columnNumber = (Double) map.get(COLUMN_NUMBER);

            LOGGER.info("Processing corrected input for get column values for keys in sheet operation");
            List<FreeForm> result = this.worksheetOperation.getRowAndColumnData(fileId, null, lookupRows, lookupColumns, getAllRecord, columnNumber, sheetName, range, null);

            LOGGER.info("Successfully retrieved column values for keys in sheet: {}", safeString(sheetName));
            return ExtensionResponseFactory.create(Map.of("Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getColumnValuesForKeysInSheetReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get column values for keys in sheet on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving column values. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getColumnValuesForKeysInSheetReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetColumnValuesForKeysConfirmation(Map<String, Object> map) {
        LOGGER.info("ColumnValuesReentryHandler.handleGetColumnValuesForKeysConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter get column values for keys data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter get column values for keys data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeysConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES_FOR_KEYS,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetColumnValuesForKeysReentry(Map<String, Object> inputMap) {
        LOGGER.info("ColumnValuesReentryHandler.handleGetColumnValuesForKeysReentry start: {}", inputMap);

        String stateId = (String) inputMap.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetColumnValuesForKeysOperationTags(
                (String) inputMap.get(TABLE_NAME), (List<String>) inputMap.get(KEY_ROWS), (List<String>) inputMap.get(LOOKUP_COLUMNS));
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeysReentry");

        try {
            String fileId = (String) inputMap.get(FILE_ID);
            String tableName = (String) inputMap.get(TABLE_NAME);
            List<String> keyRows = (List<String>) inputMap.get(KEY_ROWS);
            List<String> lookupColumns = (List<String>) inputMap.get(LOOKUP_COLUMNS);
            Boolean getAllRecord = (Boolean) inputMap.get(GET_ALL_RECORD);
            Double columnNumber = (Double) inputMap.get(COLUMN_NUMBER);

            LOGGER.info("Processing corrected input for get column values for keys operation");
            List<FreeForm> result = worksheetOperation.getRowAndColumnData(fileId, tableName, keyRows,
                    lookupColumns, getAllRecord, columnNumber, null, null, null);

            LOGGER.info("Successfully retrieved column values for keys in table: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Result", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getColumnValuesForKeysReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to get column values for keys on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving column values. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getColumnValuesForKeysReentry", startTime, telemetryTags);
        }
    }
}