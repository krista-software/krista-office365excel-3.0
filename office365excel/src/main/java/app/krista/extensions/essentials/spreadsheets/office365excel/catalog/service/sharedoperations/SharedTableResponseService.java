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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

/**
 * Helper class for creating extension responses for shared table operations
 */
@Service
public class SharedTableResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(SharedTableResponseService.class);

    private final WorkbookService workbookService;
    private final WorksheetService worksheetService;
    private final TableService tableService;
    private final WorksheetOperation worksheetOperation;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public SharedTableResponseService(WorkbookService workbookService, WorksheetService worksheetService, TableService tableService,
                                      WorksheetOperation worksheetOperation, ValidationOrchestrator validationOrchestrator,
                                      ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                      RequestContext requestContext, TelemetryHelper telemetryHelper) {
        this.workbookService = workbookService;
        this.worksheetService = worksheetService;
        this.tableService = tableService;
        this.worksheetOperation = worksheetOperation;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    public ExtensionResponse getExtensionResponse(String driveId, String fileId, String tableName, List<String> rows,
                                                  List<File> rowsFromCsvFile, Boolean allowRetry) {

        telemetryHelper.incrementCount("office365Excel.addRowsToTable");
        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddRowsToTableValidationMap(driveId, fileId, tableName, rows, rowsFromCsvFile);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableOperationTags(rows, rowsFromCsvFile);

        return handler.execute(
                "addRowsToTable",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, TABLE_NAME, tableName, ROWS, rows != null ? rows : List.of(),
                        ROWS_FROM_CSV_FILE, rowsFromCsvFile != null ? List.of(rowsFromCsvFile) : List.of()),
                SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_TABLE,
                stateId -> StateMapperUtil.addRowsToTableMetaToMap(driveId, fileId, tableName, rows, rowsFromCsvFile, stateId),
                () -> workbookService.addRowsInTable(driveId, fileId, tableName, rows, rowsFromCsvFile, null, Constants.COMMA),
                result -> {
                    LOGGER.info("Successfully added rows to table: {}", safeString(tableName));
                    return Map.of("Rows", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getTableRowsExtensionResponse(String driveId, String fileId, String tableName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getTableRows");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext,telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createTableQueryValidationMap(driveId, fileId, tableName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableQueryOperationTags(tableName);

        return handler.execute(
                "getTableRows",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, TABLE_NAME, tableName),
                SubCatalogConstants.CONFIRM_REENTER_GET_TABLE_ROWS,
                stateId -> StateMapperUtil.getTableRowsMetaToMap(driveId, fileId, tableName, stateId),
                () -> tableService.getTableData(driveId, fileId, tableName),
                result -> {
                    LOGGER.info("Successfully retrieved rows for table: {}", safeString(tableName));
                    return Map.of("Table Rows", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getDeleteColumnExtensionResponse(String driveId, String fileId, String tableName, String columnName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.deleteColumn");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createDeleteColumnValidationMap(driveId, fileId, tableName, columnName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteColumnOperationTags(tableName, columnName);

        return handler.execute(
                "deleteColumn",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, TABLE_NAME, tableName, COLUMN_NAME, columnName),
                SubCatalogConstants.CONFIRM_REENTER_DELETE_COLUMN,
                stateId -> StateMapperUtil.deleteColumnMetaToMap(driveId, fileId, tableName, columnName, stateId),
                () -> worksheetService.deleteTableColumnInSharedFile(fileId, tableName, columnName, driveId),
                result -> {
                    LOGGER.info("Successfully deleted column: {}", safeString(columnName));
                    return Map.of("Is Deleted", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getColumnValuesForKeysExtensionResponse(String driveId, String fileId, String tableName,
                                                                     List<String> keyRows, List<String> lookupColumn, Boolean getAllRecord, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeys");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createColumnValuesForKeysValidationMap(driveId, fileId, tableName, keyRows, lookupColumn);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createColumnValuesForKeysOperationTags(tableName, keyRows, lookupColumn);

        return handler.execute(
                "getColumnValuesForKeys",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, TABLE_NAME, tableName, KEY_ROWS, keyRows,
                        LOOKUP_COLUMN, lookupColumn, GET_ALL_RECORD, getAllRecord),
                SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES,
                stateId -> StateMapperUtil.getColumnValuesForKeysMetaToMap(driveId, fileId, tableName, keyRows, lookupColumn, getAllRecord, stateId),
                () -> worksheetOperation.getRowAndColumnData(fileId, tableName, keyRows, lookupColumn, getAllRecord,
                        null, null, null, driveId),
                result -> {
                    LOGGER.info("Successfully retrieved column values for keys in table: {}", safeString(tableName));
                    return Map.of("Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getVlookupInTableExtensionResponse(String driveId, String fileId, String lookupValue,
                                                                String tableName, Double columnIndex, Boolean exactMatch, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.vlookupInTable");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createVlookupInTableValidationMap(driveId, fileId, tableName, lookupValue, columnIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createVlookupInTableOperationTags(lookupValue, tableName, columnIndex);

        return handler.execute(
                "vlookupInTable",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, LOOKUP_VALUE, lookupValue, TABLE_NAME, tableName,
                        COLUMN_INDEX, columnIndex, EXACT_MATCH, exactMatch),
                SubCatalogConstants.CONFIRM_REENTER_VLOOKUP_IN_TABLE,
                stateId -> StateMapperUtil.vlookupInTableMetaToMap(driveId, fileId, lookupValue, tableName, columnIndex, exactMatch, stateId),
                () -> tableService.vlookup(driveId, fileId, lookupValue, tableName, columnIndex, exactMatch),
                result -> {
                    LOGGER.info("Successfully performed vlookup for value: {}", safeString(lookupValue));
                    return Map.of("Vlookup Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getColumnsSumExtensionResponse(String driveId, String fileId, String tableName, List<String> columns, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.columnsSum");

        AreaOperationHandler<Double> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createColumnsSumValidationMap(driveId, fileId, tableName, columns);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createColumnsSumOperationTags(tableName, columns);

        return handler.execute(
                "columnsSum",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, TABLE_NAME, tableName, COLUMNS, columns),
                SubCatalogConstants.CONFIRM_REENTER_COLUMNS_SUM,
                stateId -> StateMapperUtil.columnsSumMetaToMap(driveId, fileId, tableName, columns, stateId),
                () -> workbookService.tableColumnsSum(driveId, fileId, tableName, columns),
                result -> {
                    LOGGER.info("Successfully calculated sum for columns in table: {}", safeString(tableName));
                    return Map.of("Sum", result);
                },
                telemetryTags,
                allowRetry
        );
    }

}
