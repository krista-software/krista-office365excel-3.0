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
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import okhttp3.Response;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

/**
 * Helper class for creating extension responses for shared sheet operations
 */
@Service
public class SharedSheetResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(SharedSheetResponseService.class);

    private final WorkbookService workbookService;
    private final WorksheetService worksheetService;
    private final WorksheetOperation worksheetOperation;
    private final RestClientProvider restClientProvider;
    private final SheetFormattingService sheetFormattingService;
    private final StringUtil stringUtil;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;


    @Inject
    public SharedSheetResponseService(WorkbookService workbookService, WorksheetService worksheetService,
                                      WorksheetOperation worksheetOperation, RestClientProvider restClientProvider,
                                      SheetFormattingService sheetFormattingService, StringUtil stringUtil, ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator, RequestContext requestContext, TelemetryHelper telemetryHelper) {
        this.workbookService = workbookService;
        this.worksheetService = worksheetService;
        this.worksheetOperation = worksheetOperation;
        this.restClientProvider = restClientProvider;
        this.sheetFormattingService = sheetFormattingService;
        this.stringUtil = stringUtil;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    public ExtensionResponse getSheetExtensionResponse(String driveId, String fileId, String sheetName, List<String> rows,
                                                       List<File> rowsFromFile, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addRowsToSheet");
        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddRowsToSheetValidationMap(driveId, fileId, sheetName, rows, rowsFromFile);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createSheetOperationTags(rows, rowsFromFile);

        return handler.execute(
                "addRowsToSheet",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, ROWS, rows != null ? rows : List.of(),
                        ROWS_FROM_CSV_FILE, rowsFromFile != null ? List.of(rowsFromFile) : List.of()),
                SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_SHEET,
                stateId -> StateMapperUtil.addRowsToSheetMetaToMap(driveId, fileId, sheetName, rows, rowsFromFile, stateId),
                () -> {
                    try {
                        return worksheetService.addRows(driveId, fileId, sheetName, rows, rowsFromFile, null);
                    } catch (IOException cause) {
                        throw new RuntimeException(cause);
                    }
                },
                result -> {
                    LOGGER.info("Successfully added rows to sheet: {}", safeString(sheetName));
                    return Map.of("Sheet Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getRangeSumExtensionResponse(String driveId, String fileId, String sheetName, String range, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.rangeSum");

        AreaOperationHandler<Double> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createRangeValidationMap(driveId, fileId, sheetName, range);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags(range, sheetName);

        return handler.execute(
                "rangeSum",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, RANGE, range),
                SubCatalogConstants.CONFIRM_REENTER_RANGE_SUM,
                stateId -> StateMapperUtil.rangeSumMetaToMap(driveId, fileId, sheetName, range, stateId),
                () -> workbookService.worksheetRangeSum(driveId, fileId, sheetName, range),
                result -> {
                    LOGGER.info("Successfully calculated sum for range: {}", safeString(range));
                    return Map.of("Sum", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getSheetVlookupExtensionResponse(String driveId, String fileId, String lookupValue,
                                                              String sheetName, String range, Double columnIndex,
                                                              Boolean exactMatch, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.sheetVlookup");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);

        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createVlookupValidationMap(driveId, fileId, sheetName, range, columnIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createVlookupOperationTags(lookupValue, range, sheetName);

        return handler.execute(
                "sheetVlookup",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, LOOKUP_VALUE, lookupValue, SHEET_NAME, sheetName,
                        RANGE, range, COLUMN_INDEX, columnIndex, EXACT_MATCH, exactMatch),
                SubCatalogConstants.CONFIRM_REENTER_SHARED_SHEET_VLOOKUP,
                stateId -> StateMapperUtil.sheetVlookupMetaToMap(driveId, fileId, lookupValue, sheetName, range, columnIndex, exactMatch, stateId),
                () -> workbookService.vLookup(driveId, fileId, lookupValue, sheetName, range, columnIndex, exactMatch),
                result -> {
                    LOGGER.info("Successfully performed vlookup for value: {}", safeString(lookupValue));
                    return Map.of("Vlookup Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getRowsForRangeExtensionResponse(String driveId, String fileId, String sheetName, String range, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getRowsForRange");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createRangeValidationMap(driveId, fileId, sheetName, range);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags(range, sheetName);

        return handler.execute(
                "getRowsForRange",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, RANGE, range),
                SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE,
                stateId -> StateMapperUtil.getRowsForRangeMetaToMap(driveId, fileId, sheetName, range, stateId),
                () -> worksheetService.getRows(driveId, fileId, sheetName, range),
                result -> {
                    LOGGER.info("Successfully retrieved rows for range: {}", safeString(range));
                    return Map.of("Rows", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getSearchExtensionResponse(String driveId, String fileId, String sheetName, String dataToSearch, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.search");

        AreaOperationHandler<List<String>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createSearchValidationMap(driveId, fileId, sheetName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createSearchOperationTags(sheetName, dataToSearch);

        return handler.execute(
                "search",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, DATA_TO_SEARCH, dataToSearch),
                SubCatalogConstants.CONFIRM_REENTER_SEARCH,
                stateId -> StateMapperUtil.searchMetaToMap(driveId, fileId, sheetName, dataToSearch, stateId),
                () -> worksheetService.findValue(driveId, fileId, sheetName, dataToSearch),
                result -> {
                    LOGGER.info("Successfully searched for value: {}", safeString(dataToSearch));
                    return Map.of("Cell Ids", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getDeleteColumnWithIndexExtensionResponse(String driveId, String fileId, String sheetName, String columnName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.deleteColumnWithIndex");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createDeleteColumnWithIndexValidationMap(driveId, fileId, sheetName, columnName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteColumnWithIndexOperationTags(sheetName, columnName);

        return handler.execute(
                "deleteColumnWithIndex",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, COLUMN_NAME, columnName),
                SubCatalogConstants.CONFIRM_REENTER_DELETE_COLUMN_WITH_INDEX,
                stateId -> StateMapperUtil.deleteColumnWithIndexMetaToMap(driveId, fileId, sheetName, columnName, stateId),
                () -> {
                    //1 for the first cell of the column && 1048576 for the maximum row for any column
                    Response deleteResponse = restClientProvider.getRestClient(driveId).post(
                            fileId + Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS +
                                    columnName + 1 + Constants.COL + columnName + 1048576 + Constants.DELETE, Map.of(
                                    Constants.SHIFT, Constants.LEFT));

                    boolean success = deleteResponse.code() == 204;
                    LOGGER.info("Column deletion {}", success ? "successful" : "failed");
                    return success;
                },
                result -> Map.of("Deleted", result),
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getAddRowWithIndexExtensionResponse(String driveId, String fileId, String sheetName, List<String> rows,
                                                                 List<File> rowsFromCsv, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addRowWithIndex");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddRowWithIndexValidationMap(driveId, fileId, sheetName, rows, rowsFromCsv, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddRowWithIndexOperationTags(sheetName, rows, rowsFromCsv, rowIndex);

        return handler.execute(
                "addRowWithIndex",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, ROW_INDEX, rowIndex,
                        ROWS, rows != null ? rows : List.of(),
                        ROWS_FROM_CSV_FILE, rowsFromCsv != null ? rowsFromCsv : List.of()),
                SubCatalogConstants.CONFIRM_REENTER_ADD_ROW_WITH_INDEX,
                stateId -> StateMapperUtil.addRowWithIndexMetaToMap(driveId, fileId, sheetName, rows, rowsFromCsv, rowIndex, stateId),
                () -> {
                    try {
                        List<String> rowsStringOrCsvRows = stringUtil.getStringList(rows, rowsFromCsv);

                        restClientProvider.getRestClient(driveId).post(fileId +
                                        Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                                        + stringUtil.calculateAddressForInsert(rowIndex, rows,
                                        rowsFromCsv) + "')/insert",
                                Map.of(Constants.SHIFT, Constants.DOWN));
                        sheetFormattingService.clearRowFormatting(driveId, fileId, sheetName, rowIndex);
                        StringUtil.appendCommas(rowsStringOrCsvRows);

                        return worksheetService.addRows(driveId, fileId, sheetName, rowsStringOrCsvRows,
                                StringUtil.calculateAddressForInsert(rowIndex, rowsStringOrCsvRows));
                    } catch (IOException cause) {
                        throw new RuntimeException(cause);
                    }
                },
                result -> {
                    LOGGER.info("Successfully added row at index: {}", rowIndex);
                    return Map.of("Sheet Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getUpdateRowWithIndexExtensionResponse(String driveId, String fileId, String sheetName, String row, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.updateRowWithIndex");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createUpdateRowWithIndexValidationMap(driveId, fileId, sheetName, row, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createUpdateRowWithIndexOperationTags(sheetName, row, rowIndex);

        return handler.execute(
                "updateRowWithIndex",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, ROW, row, ROW_INDEX, rowIndex),
                SubCatalogConstants.CONFIRM_REENTER_UPDATE_ROW_WITH_INDEX,
                stateId -> StateMapperUtil.updateRowWithIndexMetaToMap(driveId, fileId, sheetName, row, rowIndex, stateId),
                () -> {
                    try {
                        return worksheetService.addRows(driveId, fileId, sheetName,
                                List.of(row), StringUtil.calculateAddressForInsert(rowIndex, List.of(row)));
                    } catch (IOException cause) {
                        throw new RuntimeException(cause);
                    }
                },
                result -> {
                    LOGGER.info("Successfully updated row at index: {}", rowIndex);
                    return Map.of("Sheet Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getCompleteRowInfoExtensionResponse(String driveId, String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getCompleteRowInfo");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetCompleteRowInfoValidationMap(driveId, fileId, sheetName, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetCompleteRowInfoOperationTags(sheetName, rowIndex);

        return handler.execute(
                "getCompleteRowInfo",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, ROW_INDEX, rowIndex),
                SubCatalogConstants.CONFIRM_REENTER_GET_COMPLETE_ROW_INFO,
                stateId -> StateMapperUtil.getCompleteRowInfoMetaToMap(driveId, fileId, sheetName, rowIndex, stateId),
                () -> worksheetService.getRow(fileId, sheetName, rowIndex, driveId),
                result -> {
                    LOGGER.info("Successfully retrieved row data for index: {}", rowIndex);
                    return Map.of("Row Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getDeleteRowWithIndexExtensionResponse(String driveId, String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.deleteRowWithIndex");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createDeleteRowWithIndexValidationMap(driveId, fileId, sheetName, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteRowWithIndexOperationTags(sheetName, rowIndex);

        return handler.execute(
                "deleteRowWithIndex",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, ROW_INDEX, rowIndex),
                SubCatalogConstants.CONFIRM_REENTER_DELETE_ROW_WITH_INDEX,
                stateId -> StateMapperUtil.deleteRowWithIndexMetaToMap(driveId, fileId, sheetName, rowIndex, stateId),
                () -> {
                    Response deleteResponse = restClientProvider.getRestClient(driveId).post(
                            fileId + Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                                    + Constants.FIRST_COLUMN + rowIndex.intValue() + Constants.COL + Constants.LAST_COLUMN
                                    + rowIndex.intValue() + Constants.DELETE, Map.of(Constants.SHIFT, Constants.UP));

                    boolean success = deleteResponse.code() == 204;
                    LOGGER.info("Row deletion {}", success ? "successful" : "failed");
                    return success;
                },
                result -> Map.of("Deleted", result),
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getColumnValuesForKeysInSheetExtensionResponse(String driveId, String fileId, String sheetName,
                                                                            String range, List<String> lookupRows, List<String> lookupColumns,
                                                                            Boolean getAllRecords, Double columnNumber, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getColumnValuesForKeysInSheet");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetColumnValuesForKeysInSheetValidationMap(driveId, fileId, sheetName, range, lookupRows, lookupColumns);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetColumnValuesForKeysInSheetOperationTags(sheetName, range, lookupRows, lookupColumns, getAllRecords, columnNumber);

        return handler.execute(
                "getColumnValuesForKeysInSheet",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, RANGE, range,
                        LOOKUP_ROWS, lookupRows, LOOKUP_COLUMNS, lookupColumns,
                        GET_ALL_RECORD, getAllRecords, COLUMN_NUMBER, columnNumber != null ? columnNumber : ""),
                SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES_IN_SHEET,
                stateId -> StateMapperUtil.getColumnValuesForKeysInSheetMetaToMap(driveId, fileId, sheetName, range, lookupRows, lookupColumns, getAllRecords, columnNumber, stateId),
                () -> worksheetOperation.getRowAndColumnData(fileId, null, lookupRows, lookupColumns,
                        getAllRecords, columnNumber, sheetName, range, driveId),
                result -> {
                    LOGGER.info("Successfully retrieved column values for keys in sheet: {}", safeString(sheetName));
                    return Map.of("Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getClearSharedSheetForRangeExtensionResponse(String driveId, String fileId, String sheetName, String cellRange, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.clearSharedSheetForRange");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createClearSharedSheetForRangeValidationMap(driveId, fileId, sheetName, cellRange);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createClearSharedSheetForRangeOperationTags(sheetName, cellRange);

        return handler.execute(
                "clearSharedSheetForRange",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, RANGE, cellRange),
                SubCatalogConstants.CONFIRM_REENTER_CLEAR_SHARED_SHEET,
                stateId -> StateMapperUtil.clearSharedSheetMetaToMap(driveId, fileId, sheetName, cellRange, stateId),
                () -> worksheetService.clearWorksheetCells(driveId, fileId, sheetName, cellRange),
                result -> {
                    LOGGER.info("Successfully cleared sheet range: {}", safeString(cellRange));
                    return Map.of("Cleared", result);
                },
                telemetryTags,
                allowRetry
        );
    }
}
