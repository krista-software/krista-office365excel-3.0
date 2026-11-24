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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.RowReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import okhttp3.Response;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

public class RowReentryHandlerImpl extends BaseReentryHandlerImpl implements RowReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowReentryHandlerImpl.class);

    private final WorksheetService worksheetService;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RestClientProvider restClientProvider;
    private final WorkbookService workbookService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public RowReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                 WorksheetService worksheetService, RestClientProvider restClientProvider, WorkbookService workbookService,
                                 TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.worksheetService = worksheetService;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.restClientProvider = restClientProvider;
        this.workbookService = workbookService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleUpdateRowWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("RowReentryHandler.handleUpdateRowWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter update row with index data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter update row with index data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.updateRowWithIndexConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_UPDATE_SHEET_ROW_WITH_INDEX,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleUpdateRowWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("RowReentryHandler.handleUpdateRowWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createUpdateRowWithIndexOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(ROW), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.updateRowWithIndexReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String row = (String) map.get(ROW);
            Double rowIndex = (Double) map.get(ROW_INDEX);

            if (row == null || row.isEmpty()) {
                LOGGER.debug("Row data not provided in reentry for fileId: {}, sheetName: {}", fileId, sheetName);
                throw new IllegalArgumentException("Row data not provided");
            }

            LOGGER.info("Processing corrected input for fileId: {}, sheetName: {}, rowIndex: {}", fileId, sheetName, rowIndex);
            List<FreeForm> result = worksheetService.addRows(null, fileId, sheetName,
                    List.of(row), StringUtil.calculateAddressForInsert(rowIndex, List.of(row)));

            LOGGER.info("Successfully updated row at index: {}", rowIndex);
            return ExtensionResponseFactory.create(Map.of("Sheet Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.updateRowWithIndexReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to update row on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while updating row at the specified index. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.updateRowWithIndexReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetCompleteRowInfoConfirmation(Map<String, Object> map) {
        LOGGER.info("RowReentryHandler.handleGetCompleteRowInfoConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter row info data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter row info data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.getCompleteRowInfoConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_COMPLETE_SHEET_ROW_INFO,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetCompleteRowInfoReentry(Map<String, Object> map) {
        LOGGER.info("RowReentryHandler.handleGetCompleteRowInfoReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetCompleteRowInfoOperationTags(
                (String) map.get(SHEET_NAME), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.getCompleteRowInfoReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            Double rowIndex = (Double) map.get(ROW_INDEX);

            if (fileId == null || fileId.isEmpty() || sheetName == null || sheetName.isEmpty() || rowIndex == null) {
                LOGGER.debug("Required parameters not provided in reentry: fileId={}, sheetName={}, rowIndex={}", fileId, sheetName, rowIndex);
                throw new IllegalArgumentException("Required parameters not provided");
            }

            FreeForm result = worksheetService.getRow(fileId, sheetName, rowIndex, null);
            LOGGER.info("Successfully retrieved row data for index: {}", rowIndex);
            return ExtensionResponseFactory.create(Map.of("Row Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getCompleteRowInfoReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get row data on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving the row data. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getCompleteRowInfoReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteRowWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("RowReentryHandler.handleDeleteRowWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter delete row with index data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter delete row with index data. Returning fetch response.");
        telemetryHelper.incrementCount("office365Excel.deleteRowWithIndexConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET_ROW_WITH_INDEX,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteRowWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("RowReentryHandler.handleDeleteRowWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteRowWithIndexOperationTags(
                (String) map.get(SHEET_NAME), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.deleteRowWithIndexReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            Double rowIndex = (Double) map.get(ROW_INDEX);

            LOGGER.info("Processing corrected input for delete row with index operation");

            // A for the column
            // XFD for the last column
            Response deleteResponse = restClientProvider.getRestClient(null).post(
                    fileId + Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                            + Constants.FIRST_COLUMN + rowIndex.intValue() + Constants.COL + Constants.LAST_COLUMN
                            + rowIndex.intValue() + Constants.DELETE, Map.of(Constants.SHIFT, Constants.UP));

            boolean success = deleteResponse.code() == 204;
            LOGGER.info("Row deletion {}", success ? "successful" : "failed");
            return ExtensionResponseFactory.create(Map.of("Deleted", success));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.deleteRowWithIndexReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to delete row on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while deleting the row. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.deleteRowWithIndexReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowsConfirmation(Map<String, Object> map) {
        LOGGER.info("AddRowsReentryHandler.handleAddRowsConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter add rows data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.addRowsConfirmation");

        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(FILE_ID, state.get(FILE_ID));
        stateMap.put(TABLE_NAME, state.get(TABLE_NAME));
        stateMap.put(ROWS, state.get(ROWS));

        Object fileObj = map.get(ROWS_FROM_CSV_FILE);
        List<File> rowsFromFileList = new ArrayList<>();

        if (fileObj instanceof List<?> files) {
            for (Object item : files) {
                if (item instanceof File) {
                    rowsFromFileList.add((File) item);
                    stateMap.put(ROWS_FROM_CSV_FILE, List.of(rowsFromFileList));
                    LOGGER.debug("Added file from list: {}", ((File) item).getFileName());
                }
            }
        } else if (fileObj instanceof File) {
            rowsFromFileList.add((File) fileObj);
            stateMap.put(ROWS_FROM_CSV_FILE, List.of(rowsFromFileList));
            LOGGER.debug("Added single file: {}", ((File) fileObj).getFileName());
        } else {
            LOGGER.debug("No valid files found in the input map");
        }

        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_IN_TABLE,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowsReentry(Map<String, Object> map) {
        LOGGER.info("AddRowsReentryHandler.handleAddRowsReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddRowsOperationTags((List<String>) map.get(ROWS),
                (List<File>) map.get(ROWS_FROM_CSV_FILE));
        telemetryHelper.incrementCount("office365Excel.addRowsReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String tableName = (String) map.get(TABLE_NAME);
            List<String> rows = (List<String>) map.get(ROWS);
            List<File> rowsFromFile = (List<File>) map.get(ROWS_FROM_CSV_FILE);

            LOGGER.info("Processing corrected input for add rows operation");
            List<FreeForm> result = workbookService.addRowsInTable(null, fileId, tableName, rows, rowsFromFile, null, Constants.COMMA);

            LOGGER.info("Successfully added rows to table: {}", tableName);
            return ExtensionResponseFactory.create(Map.of("Rows", result));
        } catch (Exception cause) {
            LOGGER.error("Add rows reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addRowsReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add rows on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding rows. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addRowsReentry", startTime, telemetryTags);
        }
    }
}