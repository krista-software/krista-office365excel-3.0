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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.ColumnReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;
import okhttp3.Response;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER;

/**
 * Implementation for column-specific reentry operations
 */
public class ColumnReentryHandlerImpl extends BaseReentryHandlerImpl  implements ColumnReentryHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(ColumnReentryHandlerImpl.class);

    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final WorksheetOperation worksheetOperation;
    private final WorkbookService workbookService;
    private final RestClientProvider restClientProvider;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public ColumnReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                    WorksheetOperation worksheetOperation, WorkbookService workbookService, RestClientProvider restClientProvider,
                                    TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.worksheetOperation = worksheetOperation;
        this.workbookService = workbookService;
        this.restClientProvider = restClientProvider;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetColumnValuesConfirmation(Map<String, Object> map) {
        LOGGER.info("ColumnReentryHandler.handleGetColumnValuesConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter column values data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter column values data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.getColumnValuesConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetColumnValuesReentry(Map<String, Object> map) {
        LOGGER.info("ColumnReentryHandler.handleGetColumnValuesReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createColumnValuesForKeysOperationTags(
                (String) map.get(TABLE_NAME), (List<String>) map.get(KEY_ROWS), (List<String>) map.get(LOOKUP_COLUMN));
        telemetryHelper.incrementCount("office365Excel.getColumnValuesReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);
            List<String> keyRows = (List<String>) map.get(KEY_ROWS);
            List<String> lookupColumn = (List<String>) map.get(LOOKUP_COLUMN);
            Boolean getAllRecord = (Boolean) map.get(GET_ALL_RECORD);

            LOGGER.info("Processing corrected input for get column values operation");
            List<FreeForm> result = this.worksheetOperation.getRowAndColumnData(fileId, tableName, keyRows, lookupColumn, getAllRecord,
                    null, null, null, driveId);

            LOGGER.info("Successfully retrieved column values for keys in table: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Result", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getColumnValuesReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get column values on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving column values. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getColumnValuesReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleColumnsSumConfirmation(Map<String, Object> map) {
        LOGGER.info("ColumnReentryHandler.handleColumnsSumConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter columns sum data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter columns sum data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.columnsSumConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_COLUMNS_SUM,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleColumnsSumReentry(Map<String, Object> map) {
        LOGGER.info("ColumnReentryHandler.handleColumnsSumReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createColumnsSumOperationTags(
                (String) map.get(TABLE_NAME), (List<String>) map.get(COLUMNS));
        telemetryHelper.incrementCount("office365Excel.columnsSumReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);
            List<String> columns = (List<String>) map.get(COLUMNS);

            LOGGER.info("Processing corrected input for columns sum operation");
            Double result = this.workbookService.tableColumnsSum(driveId, fileId, tableName, columns);

            LOGGER.info("Successfully calculated sum for columns in table: {}", safeString(tableName));
            return ExtensionResponseFactory.create(Map.of("Sum", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.columnsSumReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to calculate sum on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while calculating sum. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.columnsSumReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteColumnWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("ColumnReentryHandler.handleDeleteColumnWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter delete column with index data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter delete column with index data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.deleteColumnWithIndexConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteColumnWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("ColumnReentryHandler.handleDeleteColumnWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteColumnWithIndexOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(COLUMN_NAME));
        telemetryHelper.incrementCount("office365Excel.deleteColumnWithIndexReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String columnName = (String) map.get(COLUMN_NAME);

            LOGGER.info("Processing corrected input for delete column with index operation");

            Response deleteResponse = restClientProvider.getRestClient(null).post(
                    fileId + Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS +
                            columnName + 1 + Constants.COL + columnName + 1048576 + Constants.DELETE, Map.of(
                            Constants.SHIFT, Constants.LEFT));

            boolean success = deleteResponse.code() == 204;
            LOGGER.info("Column deletion {}", success ? "successful" : "failed");
            return ExtensionResponseFactory.create(Map.of("Deleted", success));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.deleteColumnWithIndexReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to delete column on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while deleting the column. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.deleteColumnWithIndexReentry", startTime, telemetryTags);
        }
    }
}
