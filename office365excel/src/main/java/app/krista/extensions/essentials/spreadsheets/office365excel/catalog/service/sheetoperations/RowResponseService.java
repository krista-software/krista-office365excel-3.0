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

@Service
public class RowResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowResponseService.class);

    private final WorksheetService worksheetService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    private final RestClientProvider restClientProvider;
    private final StringUtil stringUtil;
    private final SheetFormattingService sheetFormattingService;

    @Inject
    public RowResponseService(WorksheetService worksheetService, ValidationOrchestrator validationOrchestrator,
                              ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                              RequestContext requestContext, TelemetryHelper telemetryHelper, RestClientProvider restClientProvider,
                              StringUtil stringUtil, SheetFormattingService sheetFormattingService) {
        this.worksheetService = worksheetService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
        this.restClientProvider = restClientProvider;
        this.stringUtil = stringUtil;
        this.sheetFormattingService = sheetFormattingService;
    }

    public ExtensionResponse getAddRowsExtensionResponse(String fileId, String sheetName, List<String> rows, List<File> rowsFromCsv, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addRows");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddRowsValidationMap(fileId, sheetName, rows, rowsFromCsv);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddRowsOperationTags(rows, rowsFromCsv);

        return handler.execute(
                "addRows",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, ROWS, rows != null ? rows : List.of(),
                        ROWS_FROM_CSV_FILE, rowsFromCsv != null ? rowsFromCsv : List.of()),
                SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS,
                stateId -> StateMapperUtil.addRowsMetaToMap(fileId, sheetName, rows, rowsFromCsv, stateId),
                () -> {
                    try {
                        return worksheetService.addRows(null, fileId, sheetName, rows, rowsFromCsv, null);
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

    public ExtensionResponse getAddRowWithIndexExtensionResponse(String fileId, String sheetName, List<String> rows, List<File> rowsFromCsv, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addRowWithIndex");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddRowWithIndexValidationMapForSheet(fileId, sheetName, rows, rowsFromCsv, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddRowWithIndexOperationTags(sheetName, rows, rowsFromCsv, rowIndex);

        return handler.execute(
                "addRowWithIndex",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, ROW_INDEX, rowIndex,
                        ROWS, rows != null ? rows : List.of(),
                        ROWS_FROM_CSV_FILE, rowsFromCsv != null ? rowsFromCsv : List.of()),
                SubCatalogConstants.CONFIRM_REENTER_ADD_SHEET_ROW_WITH_INDEX,
                stateId -> StateMapperUtil.addRowWithIndexMetaToMap(fileId, sheetName, rows, rowsFromCsv, rowIndex, stateId),
                () -> addRowsAtIndex(fileId, sheetName, rows, rowsFromCsv, rowIndex),
                result -> {
                    LOGGER.info("Successfully added row at index: {}", rowIndex);
                    return Map.of("Sheet Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    private List<FreeForm> addRowsAtIndex(String fileId, String sheetName, List<String> rows, List<File> rowsFromCsv, Double rowIndex) {
        try {
            List<String> rowsStringOrCsvRows = stringUtil.getStringList(rows, rowsFromCsv);

            restClientProvider.getRestClient(null).post(fileId +
                            Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                            + stringUtil.calculateAddressForInsert(rowIndex, rows, rowsFromCsv) + "')/insert",
                    Map.of(Constants.SHIFT, Constants.DOWN));

            sheetFormattingService.clearRowFormatting(fileId, sheetName, rowIndex, rowsStringOrCsvRows.size());
            StringUtil.appendCommas(rowsStringOrCsvRows);

            return worksheetService.addRows(null, fileId, sheetName, rowsStringOrCsvRows,
                    StringUtil.calculateAddressForInsert(rowIndex, rowsStringOrCsvRows));
        } catch (IOException cause) {
            throw new RuntimeException(cause);
        }
    }

    public ExtensionResponse getGetCompleteRowInfoExtensionResponse(String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getCompleteRowInfo");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetCompleteRowInfoValidationMapForSheet(fileId, sheetName, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetCompleteRowInfoOperationTags(sheetName, rowIndex);

        return handler.execute(
                "getCompleteRowInfo",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, ROW_INDEX, rowIndex),
                SubCatalogConstants.CONFIRM_REENTER_GET_COMPLETE_SHEET_ROW_INFO,
                stateId -> StateMapperUtil.getCompleteRowInfoMetaToMap(null, fileId, sheetName, rowIndex, stateId),
                () -> worksheetService.getRow(fileId, sheetName, rowIndex, null),
                result -> {
                    LOGGER.info("Successfully retrieved row data for index: {}", rowIndex);
                    return Map.of("Row Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getGetRowsForRangeExtensionResponse(String fileID, String sheetName, String range, Boolean allowRetry) {
        return getGetRowsForRangeExtensionResponse(null, fileID, sheetName, range, allowRetry);
    }

    public ExtensionResponse getGetRowsForRangeExtensionResponse(String driveId, String fileID, String sheetName, String range, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getRowsForRange");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createRangeValidationMapForSheet(fileID, sheetName, range);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags(range, sheetName);

        return handler.execute(
                "getRowsForRange",
                validationMap,
                Map.of(FILE_ID, fileID, SHEET_NAME, sheetName, RANGE, range),
                SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET,
                stateId -> StateMapperUtil.getRowsForRangeMetaToMap(driveId, fileID, sheetName, range, stateId),
                () -> worksheetService.getRows(driveId, fileID, sheetName, range),
                result -> {
                    LOGGER.info("Successfully retrieved rows for range: {}", safeString(range));
                    return Map.of("Rows", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getDeleteRowWithIndexExtensionResponse(String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.deleteRowWithIndex");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createDeleteRowWithIndexValidationMapForSheet(fileId, sheetName, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteRowWithIndexOperationTags(sheetName, rowIndex);

        return handler.execute(
                "deleteRowWithIndex",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, ROW_INDEX, rowIndex),
                SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET_ROW_WITH_INDEX,
                stateId -> StateMapperUtil.deleteRowWithIndexMetaToMap(null, fileId, sheetName, rowIndex, stateId),
                () -> deleteRowFromSheet(fileId, sheetName, rowIndex),
                result -> {
                    LOGGER.info("Successfully deleted row at index: {}", rowIndex);
                    return Map.of("Deleted", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    private boolean deleteRowFromSheet(String fileId, String sheetName, Double rowIndex) {
        // A for the column & XFD for the last column
        Response deleteResponse = restClientProvider.getRestClient(null).post(
                fileId + Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                        + Constants.FIRST_COLUMN + rowIndex.intValue() + Constants.COL + Constants.LAST_COLUMN
                        + rowIndex.intValue() + Constants.DELETE, Map.of(Constants.SHIFT, Constants.UP));

        boolean success = deleteResponse.code() == 204;
        LOGGER.info("Row deletion {}", success ? "successful" : "failed");
        return success;
    }

    public ExtensionResponse getUpdateRowWithIndexExtensionResponse(String fileId, String sheetName, String row, Double rowIndex, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.updateRowWithIndex");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createUpdateRowWithIndexValidationMapForSheet(fileId, sheetName, row, rowIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createUpdateRowWithIndexOperationTags(sheetName, row, rowIndex);

        return handler.execute(
                "updateRowWithIndex",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, ROW, row, ROW_INDEX, rowIndex),
                SubCatalogConstants.CONFIRM_REENTER_UPDATE_SHEET_ROW_WITH_INDEX,
                stateId -> StateMapperUtil.updateRowWithIndexMetaToMap(null, fileId, sheetName, row, rowIndex, stateId),
                () -> {
                    try {
                        return worksheetService.addRows(null, fileId, sheetName,
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

}
