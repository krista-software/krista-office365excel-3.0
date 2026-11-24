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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.TableReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;

import javax.inject.Inject;
import java.util.*;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

/**
 * Implementation for table-specific reentry operations
 */
public class TableReentryHandlerImpl extends BaseReentryHandlerImpl implements TableReentryHandler {

    private final WorkbookService workbookService;
    private final TableService tableService;
    private final WorksheetService worksheetService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public TableReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                   WorkbookService workbookService, TableService tableService, WorksheetService worksheetService,
                                   TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.workbookService = workbookService;
        this.tableService = tableService;
        this.worksheetService = worksheetService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleConfirmation(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter table data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter table data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.tableConfirmation");

        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(DRIVE_ID, map.get(DRIVE_ID));
        stateMap.put(FILE_ID, map.get(FILE_ID));
        stateMap.put(TABLE_NAME, map.get(TABLE_NAME));
        stateMap.put(ROWS, state.get(ROWS));

        Object fileObj = map.get(ROWS_FROM_CSV_FILE);
        List<File> rowsFromFileList = new ArrayList<>();

        if (fileObj instanceof List<?> files) {
            for (Object item : files) {
                if (item instanceof File) {
                    rowsFromFileList.add((File) item);
                    LOGGER.debug("Added file from list: {}", ((File) item).getFileName());
                }
            }
        } else if (fileObj instanceof File) {
            rowsFromFileList.add((File) fileObj);
            LOGGER.debug("Added single file: {}", ((File) fileObj).getFileName());
        } else {
            LOGGER.debug("No valid files found in the input map");
        }

        if (!rowsFromFileList.isEmpty()) {
            stateMap.put(ROWS_FROM_CSV_FILE, rowsFromFileList);
        } else {
            stateMap.put(ROWS_FROM_CSV_FILE, state.get(ROWS_FROM_CSV_FILE));
        }

        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_ADD_ROWS_TO_TABLE,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReentry(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableOperationTags(
                (List<String>) map.get(ROWS), (List<File>) map.get(ROWS_FROM_CSV_FILE));
        telemetryHelper.incrementCount("office365Excel.tableReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);
            List<String> rows = (List<String>) map.get(ROWS);
            List<File> rowsFromCsvFile = (List<File>) map.get(ROWS_FROM_CSV_FILE);

            if (!Worksheets.isRowDataAvailable(rows, rowsFromCsvFile)) {
                LOGGER.debug("Row data not provided in reentry for driveId: {}, fileId: {}, tableName: {}",
                        driveId, fileId, tableName);
                throw new IllegalArgumentException("Row data not provided");
            }

            LOGGER.info("Processing corrected input for driveId: {}, fileId: {}, tableName: {}", driveId, fileId, tableName);
            List<FreeForm> result = this.workbookService.addRowsInTable(driveId, fileId, tableName, rows, rowsFromCsvFile, null, Constants.COMMA);

            LOGGER.info("Successfully added rows to table: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Rows", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.tableReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause, "Failed to add rows on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding rows to the table. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.tableReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetTableRowsConfirmation(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleGetTableRowsConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter get table rows data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        String driveId = (String) map.get(DRIVE_ID);
        String fileId = (String) map.get(FILE_ID);
        String tableName = (String) map.get(TABLE_NAME);

        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(DRIVE_ID, driveId);
        stateMap.put(FILE_ID, fileId);
        stateMap.put(TABLE_NAME, tableName);

        LOGGER.info("User chose to re-enter get table rows data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.getTableRowsConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_TABLE_ROWS,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetTableRowsReentry(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleGetTableRowsReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableQueryOperationTags((String) map.get(TABLE_NAME));
        telemetryHelper.incrementCount("office365Excel.getTableRowsReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);

            LOGGER.info("Processing corrected input for driveId: {}, fileId: {}, tableName: {}", driveId, fileId, tableName);

            List<FreeForm> result = this.tableService.getTableData(driveId, fileId, tableName);

            LOGGER.info("Successfully retrieved rows from table: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Rows", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getTableRowsReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause, "Failed to get table rows on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving rows from the table. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getTableRowsReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleVlookupInTableConfirmation(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleVlookupInTableConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter vlookup data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.vlookupInTableConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_VLOOKUP_IN_TABLE,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleVlookupInTableReentry(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleVlookupInTableReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createVlookupInTableOperationTags(
                (String) map.get(LOOKUP_VALUE), (String) map.get(TABLE_NAME), (Double) map.get(COLUMN_INDEX));
        telemetryHelper.incrementCount("office365Excel.vlomokupInTableReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String lookupValue = (String) map.get(LOOKUP_VALUE);
            String tableName = (String) map.get(TABLE_NAME);
            Double columnIndex = (Double) map.get(COLUMN_INDEX);
            Boolean exactMatch = (Boolean) map.get(EXACT_MATCH);

            LOGGER.info("Processing corrected input for vlookup in table operation");
            FreeForm result = this.tableService.vlookup(driveId, fileId, lookupValue, tableName, columnIndex, exactMatch);

            LOGGER.info("Successfully performed vlookup for value: {}", lookupValue);
            return ExtensionResponseFactory.create(Map.of("Vlookup Result", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.vlookupInTableReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to perform vlookup on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while performing vlookup. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.vlookupInTableReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteColumnConfirmation(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleDeleteColumnConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter delete column data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.deleteColumnConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_DELETE_COLUMN,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteColumnReentry(Map<String, Object> map) {
        LOGGER.info("TableReentryHandler.handleDeleteColumnReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteColumnOperationTags(
                (String) map.get(TABLE_NAME), (String) map.get(COLUMN_NAME));
        telemetryHelper.incrementCount("office365Excel.deleteColumnReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);
            String columnName = (String) map.get(COLUMN_NAME);

            LOGGER.info("Processing corrected input for delete column operation");

            Boolean result = worksheetService.deleteTableColumnInSharedFile(fileId, tableName, columnName, driveId);

            LOGGER.info("Successfully deleted column: {}", columnName);
            return ExtensionResponseFactory.create(Map.of("Is Deleted", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.deleteColumnReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to delete column on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while deleting the column. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.deleteColumnReentry", startTime, telemetryTags);
        }
    }
}
