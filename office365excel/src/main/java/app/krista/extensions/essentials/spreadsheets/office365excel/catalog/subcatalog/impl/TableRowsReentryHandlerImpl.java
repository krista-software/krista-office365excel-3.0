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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.TableRowsReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

public class TableRowsReentryHandlerImpl extends BaseReentryHandlerImpl implements TableRowsReentryHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(TableRowsReentryHandlerImpl.class);

    private final TableService tableService;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final WorkbookService workbookService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public TableRowsReentryHandlerImpl(TableService tableService, ErrorHandlingStateManager internalStateManager,
                                       ExtensionResponseGenerator responseGenerator, WorkbookService workbookService,
                                       TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.tableService = tableService;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.workbookService = workbookService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleConfirmation(Map<String, Object> map) {
        LOGGER.info("GetRowsReentryHandler.handleConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter get rows data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.getRowsConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_GET_ROWS,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReentry(Map<String, Object> map) {
        LOGGER.info("GetRowsReentryHandler.handleReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableQueryOperationTags((String) map.get(TABLE_NAME));
        telemetryHelper.incrementCount("office365Excel.getRowsReentry");

        try {
            String workbookId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);

            LOGGER.info("Processing corrected input for workbookId: {}, tableName: {}", workbookId, tableName);
            List<FreeForm> result = tableService.getTableData(null, workbookId, tableName);

            LOGGER.info("Successfully retrieved rows for table: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Table Rows", result));
        } catch (Exception cause) {
            LOGGER.error("Error in handleReentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getRowsReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get table rows on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving rows from the table. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getRowsReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowsToTableConfirmation(Map<String, Object> map) {
        LOGGER.info("TableWithIndexReentryHandler.handleConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter table with index data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.addRowsToTableWithIndexConfirmation");

        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(DRIVE_ID, state.get(DRIVE_ID));
        stateMap.put(FILE_ID, state.get(FILE_ID));
        stateMap.put(TABLE_NAME, state.get(TABLE_NAME));
        stateMap.put(ROWS, state.get(ROWS));
        stateMap.put(ROW_INDEX, state.get(ROW_INDEX));
        stateMap.put(SEPARATOR, state.get(SEPARATOR));
        SubCatalogUtil.processAndUpdateFiles(map, stateMap, state, ROWS_FROM_CSV_FILE);

        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowsToTableReentry(Map<String, Object> map) {
        LOGGER.info("TableWithIndexReentryHandler.handleReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableOperationTags(
                (List<String>) map.get(ROWS), (List<File>) map.get(ROWS_FROM_CSV_FILE));
        telemetryHelper.incrementCount("office365Excel.addRowsToTableWithIndexReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);
            List<String> rows = (List<String>) map.get(ROWS);
            List<File> rowsFromCsvFile = (List<File>) map.get(ROWS_FROM_CSV_FILE);
            Double rowIndex = (Double) map.get(ROW_INDEX);
            String separator = (String) map.get(SEPARATOR);

            LOGGER.info("Processing corrected input for driveId: {}, fileId: {}, tableName: {}, rowIndex: {}",
                    driveId, fileId, tableName, rowIndex);
            List<FreeForm> result = workbookService.addRowsInTable(driveId, fileId, tableName, rows,
                    rowsFromCsvFile, rowIndex.intValue() - 1, separator == null ? Constants.COMMA : separator);

            LOGGER.info("Successfully added rows to table: {} at index: {}", tableName, rowIndex);
            return ExtensionResponseFactory.create(Map.of("Total Rows", result));
        } catch (Exception cause) {
            LOGGER.error("Error in handleReentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addRowsToTableWithIndexReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add rows on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding rows. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addRowsToTableWithIndexReentry", startTime, telemetryTags);
        }
    }

}
