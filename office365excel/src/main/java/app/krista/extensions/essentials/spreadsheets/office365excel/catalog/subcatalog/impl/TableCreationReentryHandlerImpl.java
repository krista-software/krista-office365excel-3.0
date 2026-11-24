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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.ExtensionResponseFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.RemediationActionFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.TableCreationReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER;

/**
 * Implementation for table creation reentry operations
 */
public class TableCreationReentryHandlerImpl extends BaseReentryHandlerImpl implements TableCreationReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(TableCreationReentryHandlerImpl.class);
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final TableService tableService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public TableCreationReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                           TableService tableService, TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.tableService = tableService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleCreateTableForRangeConfirmation(Map<String, Object> map) {
        LOGGER.info("TableCreationReentryHandler.handleCreateTableForRangeConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter create table for range data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter create table for range data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.createTableForRangeConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_CREATE_TABLE_FOR_RANGE,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleCreateTableForRangeReentry(Map<String, Object> inputMap) {
        LOGGER.info("TableCreationReentryHandler.handleCreateTableForRangeReentry start: {}", inputMap);

        String stateId = (String) inputMap.get(STATE_ID);
        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags(
                (String) inputMap.get(SHEET_NAME), (String) inputMap.get(RANGE));
        telemetryHelper.incrementCount("office365Excel.createTableForRangeReentry");

        try {
            String fileId = (String) inputMap.get(FILE_ID);
            String sheetName = (String) inputMap.get(SHEET_NAME);
            String range = (String) inputMap.get(RANGE);
            String driveId = (String) inputMap.get(DRIVE_ID);

            LOGGER.info("Processing corrected input for create table for range operation");
            Table result = tableService.createWorkbookTable(fileId, sheetName, range, driveId);

            LOGGER.info("Successfully created table for range: {}", range);
            return ExtensionResponseFactory.create(Map.of("Table", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.createTableForRangeReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to create table for range on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while creating the table. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.createTableForRangeReentry", startTime, telemetryTags);
        }
    }
}