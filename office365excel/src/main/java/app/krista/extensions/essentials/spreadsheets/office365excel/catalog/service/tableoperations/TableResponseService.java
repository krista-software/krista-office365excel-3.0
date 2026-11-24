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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.tableoperations;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
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

@Service
public class TableResponseService {
    
    private static final Logger LOGGER = LoggerFactory.getLogger(TableResponseService.class);
    
    private final WorkbookService workbookService;
    private final TableService tableService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    private final WorksheetService worksheetService;

    @Inject
    public TableResponseService(WorkbookService workbookService, TableService tableService,
                                ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager,
                                ExtensionResponseGenerator responseGenerator, RequestContext requestContext,
                                TelemetryHelper telemetryHelper, WorksheetService worksheetService) {
        this.workbookService = workbookService;
        this.tableService = tableService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
        this.worksheetService = worksheetService;
    }
    
    public ExtensionResponse getAddRowsToTableExtensionResponse(String fileID, String driveID, String tableName,
                                                               List<String> rows, List<File> rowsFromCsv,
                                                               Double index, String separator, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addRowsToTableWithIndex");
        
        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddRowsToTableValidationMapForTable(fileID, tableName, rows, rowsFromCsv, index);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableWithIndexOperationTags(tableName, rows, rowsFromCsv, index);

        return handler.execute(
                "addRowsToTable",
                validationMap,
                Map.of(FILE_ID, fileID, DRIVE_ID, driveID, TABLE_NAME, tableName, ROWS, rows != null ? rows : List.of(),
                        ROWS_FROM_CSV_FILE, rowsFromCsv != null ? List.of(rowsFromCsv) : List.of(), ROW_INDEX, index, SEPARATOR, separator),
                SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX,
                stateId -> StateMapperUtil.addRowsToTableWithIndexMetaToMap(driveID, fileID, tableName, rows, rowsFromCsv, index, separator, stateId),
                () -> workbookService.addRowsInTable(driveID, fileID, tableName, rows, rowsFromCsv,
                        index.intValue() - 1, separator),
                result -> {
                    LOGGER.info("Successfully added rows to table: {} at index: {}", safeString(tableName), index);
                    return Map.of("Total Rows", result);
                },
                telemetryTags,
                allowRetry
        );
    }
    
    public ExtensionResponse getTableRowsExtensionResponse(String workbookID, String tableName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getTableRows");
        
        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createTableQueryValidationMapForTable(workbookID, tableName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableQueryOperationTags(tableName);

        return handler.execute(
                "getTableRows",
                validationMap,
                Map.of(FILE_ID, workbookID, TABLE_NAME, tableName),
                SubCatalogConstants.CONFIRM_REENTER_GET_ROWS,
                stateId -> StateMapperUtil.getRowsMetaToMap(workbookID, tableName, stateId),
                () -> this.tableService.getTableData(null, workbookID, tableName),
                result -> {
                    LOGGER.info("Successfully retrieved rows for table: {}", safeString(tableName));
                    return Map.of("Table Rows", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getAddRowsExtensionResponse(String fileID, String tableName, List<String> rows, List<File> rowsFromCsv, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addRowsInTable");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddRowsInTableValidationMap(fileID, tableName, rows, rowsFromCsv);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddRowsInTableOperationTags(tableName, rows, rowsFromCsv);

        return handler.execute(
                "addRowsInTable",
                validationMap,
                Map.of(FILE_ID, fileID, TABLE_NAME, tableName, ROWS, rows != null ? rows : List.of(),
                        ROWS_FROM_CSV_FILE, rowsFromCsv != null ? List.of(rowsFromCsv) : List.of()),
                SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_IN_TABLE,
                stateId -> StateMapperUtil.addTableRowsMetaToMap(fileID, tableName, rows, rowsFromCsv, stateId),
                () -> workbookService.addRowsInTable(null, fileID, tableName, rows, rowsFromCsv, null, Constants.COMMA),
                result -> {
                    LOGGER.info("Successfully added rows to table: {}", safeString(tableName));
                    return Map.of("Rows", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getDeleteTableColumnExtensionResponse(String fileID, String tableName, String columnName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.deleteTableColumn");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createDeleteTableColumnValidationMap(fileID, tableName, columnName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteTableColumnOperationTags(tableName, columnName);

        return handler.execute(
                "deleteTableColumn",
                validationMap,
                Map.of(FILE_ID, fileID, TABLE_NAME, tableName, COLUMN_NAME, columnName),
                SubCatalogConstants.CONFIRM_REENTER_DELETE_TABLE_COLUMN,
                stateId -> StateMapperUtil.deleteTableColumnMetaToMap(fileID, tableName, columnName, stateId),
                () -> worksheetService.deleteTableColumn(fileID, tableName, columnName),
                result -> {
                    LOGGER.info("Successfully processed delete column for table: {}", safeString(tableName));
                    return Map.of("Is Deleted", result);
                },
                telemetryTags,
                allowRetry
        );
    }
}