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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sheetoperations;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.FreeForm;
import okhttp3.Response;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class ColumnResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(ColumnResponseService.class);

    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    private final RestClientProvider restClientProvider;
    private final WorksheetOperation worksheetOperation;

    @Inject
    public ColumnResponseService(ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager,
                                 ExtensionResponseGenerator responseGenerator, RequestContext requestContext, TelemetryHelper telemetryHelper,
                                 RestClientProvider restClientProvider, WorksheetOperation worksheetOperation) {
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
        this.restClientProvider = restClientProvider;
        this.worksheetOperation = worksheetOperation;
    }

    public ExtensionResponse getDeleteColumnWithIndexExtensionResponse(String fileId, String sheetName, String columnName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.deleteColumnWithIndex");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createDeleteColumnWithIndexValidationMapForSheet(fileId, sheetName, columnName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteColumnWithIndexOperationTags(sheetName, columnName);

        return handler.execute(
                "deleteColumnWithIndex",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, COLUMN_NAME, columnName),
                SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX,
                stateId -> StateMapperUtil.deleteColumnWithIndexMetaToMap(null, fileId, sheetName, columnName, stateId),
                () -> deleteColumnFromSheet(fileId, sheetName, columnName),
                result -> {
                    LOGGER.info("Successfully deleted column: {}", safeString(columnName));
                    return Map.of("Deleted", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    private boolean deleteColumnFromSheet(String fileId, String sheetName, String columnName) {
        //1 for the first cell of the column && 1048576 for the maximum row for any column
        Response deleteResponse = restClientProvider.getRestClient(null).post(
                fileId + Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS +
                        columnName + 1 + Constants.COL + columnName + 1048576 + Constants.DELETE, Map.of(
                        Constants.SHIFT, Constants.LEFT));

        boolean success = deleteResponse.code() == 204;
        LOGGER.info("Column deletion {}", success ? "successful" : "failed");
        return success;
    }

    public ExtensionResponse getGetColumnValuesForKeysInSheetExtensionResponse(String fileId, String sheetName, String range, List<String> keyRows, List<String> keyColumns, Boolean getAllRecords, Double columnNumber, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeysInSheet");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetColumnValuesForKeysInSheetValidationMapForSheet(fileId, sheetName, range, keyRows, keyColumns);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetColumnValuesForKeysInSheetOperationTags(sheetName, range, keyRows, keyColumns, getAllRecords, columnNumber);

        return handler.execute(
                "getColumnValuesForKeysInSheet",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, RANGE, range, LOOKUP_ROWS, keyRows, LOOKUP_COLUMNS, keyColumns,
                        GET_ALL_RECORD, getAllRecords, COLUMN_NUMBER, columnNumber != null ? columnNumber : ""),
                SubCatalogConstants.CONFIRM_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET,
                stateId -> StateMapperUtil.getColumnValuesForKeysInSheetMetaToMap(null, fileId, sheetName, range, keyRows, keyColumns, getAllRecords, columnNumber, stateId),
                () -> this.worksheetOperation.getRowAndColumnData(fileId, null, keyRows, keyColumns, getAllRecords, columnNumber, sheetName, range, null),
                result -> {
                    LOGGER.info("Successfully retrieved column values for keys in sheet: {}", safeString(sheetName));
                    return Map.of("Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

}
