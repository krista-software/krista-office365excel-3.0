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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.DeleteColumnReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER;

/**
 * Implementation for delete column reentry operations
 */
public class DeleteColumnReentryHandlerImpl extends BaseReentryHandlerImpl implements DeleteColumnReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(DeleteColumnReentryHandlerImpl.class);

    private final WorksheetService worksheetService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public DeleteColumnReentryHandlerImpl(ErrorHandlingStateManager internalStateManager,
                                          ExtensionResponseGenerator responseGenerator,
                                          WorksheetService worksheetService, TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.worksheetService = worksheetService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteTableColumnConfirmation(Map<String, Object> map) {
        LOGGER.info("TableColumnReentryHandler.handleDeleteTableColumnConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter delete table column data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter delete table column data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.deleteTableColumnConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_DELETE_TABLE_COLUMN,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteTableColumnReentry(Map<String, Object> inputMap) {
        LOGGER.info("TableColumnReentryHandler.handleDeleteTableColumnReentry start: {}", inputMap);
        String stateId = (String) inputMap.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteTableColumnOperationTags(
                (String) inputMap.get(TABLE_NAME), (String) inputMap.get(COLUMN_NAME));
        telemetryHelper.incrementCount("office365Excel.deleteTableColumnReentry");

        try {
            String fileId = (String) inputMap.get(FILE_ID);
            String tableName = (String) inputMap.get(TABLE_NAME);
            String columnName = (String) inputMap.get(COLUMN_NAME);

            LOGGER.info("Processing corrected input for delete table column operation");
            boolean result = this.worksheetService.deleteTableColumn(fileId, tableName, columnName);

            LOGGER.info("Successfully processed delete table column for: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Is Deleted", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.deleteTableColumnReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to delete table column on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while deleting the column. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.deleteTableColumnReentry", startTime, telemetryTags);
        }
    }
}