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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.FreeForm;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class TableCreationResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(TableCreationResponseService.class);

    private final TableService tableService;
    private final WorksheetOperation worksheetOperation;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public TableCreationResponseService(TableService tableService, WorksheetOperation worksheetOperation,
                                        ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager,
                                        ExtensionResponseGenerator responseGenerator, RequestContext requestContext,
                                        TelemetryHelper telemetryHelper) {
        this.tableService = tableService;
        this.worksheetOperation = worksheetOperation;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    public ExtensionResponse getCreateTableForRangeExtensionResponse(String fileId, String sheetName, String range, String driveId, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.createTableForRange");

        AreaOperationHandler<Table> handler = new AreaOperationHandler<>(validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createTableCreationValidationMap(fileId, sheetName, range);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableCreationOperationTags(sheetName, range);

        return handler.execute(
                "createTableForRange",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, RANGE, range, DRIVE_ID, driveId != null ? driveId : ""),
                SubCatalogConstants.CONFIRM_REENTER_CREATE_TABLE_FOR_RANGE,
                stateId -> StateMapperUtil.createTableForRangeMetaToMap(fileId, sheetName, range, driveId, stateId),
                () -> this.tableService.createWorkbookTable(fileId, sheetName, range, driveId),
                result -> {
                    LOGGER.info("Successfully created table for range: {}", safeString(range));
                    return Map.of("Table", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getColumnValuesForKeysExtensionResponse(String fileID, String tableName, List<String> keyRows,
                                                                     List<String> lookupColumns, Boolean getAllRecord, Double columnNumber, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeys");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createColumnValuesForKeysValidationMapForTable(fileID, tableName, keyRows, lookupColumns);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createColumnValuesForKeysOperationTags(tableName, keyRows, lookupColumns);

        return handler.execute(
                "getColumnValuesForKeys",
                validationMap,
                Map.of(FILE_ID, fileID, TABLE_NAME, tableName, KEY_ROWS, keyRows != null ? keyRows : List.of(), LOOKUP_COLUMNS,
                        lookupColumns != null ? lookupColumns : List.of(), GET_ALL_RECORD, getAllRecord),
                SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES_FOR_KEYS,
                stateId -> StateMapperUtil.getTableColumnValuesForKeysMetaToMap(fileID, tableName, keyRows, lookupColumns, getAllRecord, columnNumber, stateId),
                () -> this.worksheetOperation.getRowAndColumnData(fileID, tableName, keyRows, lookupColumns, getAllRecord,
                        columnNumber, null, null, null),
                result -> {
                    LOGGER.info("Successfully retrieved column values for keys in table: {}", safeString(tableName));
                    return Map.of("Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }
}