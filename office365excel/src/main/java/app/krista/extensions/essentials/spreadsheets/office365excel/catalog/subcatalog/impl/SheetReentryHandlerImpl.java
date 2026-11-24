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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.SheetReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.SheetFormattingService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import okhttp3.Response;

import javax.inject.Inject;
import java.util.*;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

/**
 * Implementation for sheet-specific reentry operations
 */
public class SheetReentryHandlerImpl extends BaseReentryHandlerImpl implements SheetReentryHandler {

    private final WorksheetService worksheetService;
    private final RestClientProvider restClientProvider;
    private final StringUtil stringUtil;
    private final WorksheetOperation worksheetOperation;
    private final SheetFormattingService sheetFormattingService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public SheetReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                   WorksheetService worksheetService, RestClientProvider restClientProvider, StringUtil stringUtil,
                                   WorksheetOperation worksheetOperation, SheetFormattingService sheetFormattingService,
                                   TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.worksheetService = worksheetService;
        this.restClientProvider = restClientProvider;
        this.stringUtil = stringUtil;
        this.worksheetOperation = worksheetOperation;
        this.sheetFormattingService = sheetFormattingService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter sheet data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults, null, Map.of());
        }

        LOGGER.info("User chose to re-enter sheet data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.sheetReentryConfirmation");

        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(DRIVE_ID, state.get(DRIVE_ID));
        stateMap.put(FILE_ID, state.get(FILE_ID));
        stateMap.put(SHEET_NAME, state.get(SHEET_NAME));
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
                HANDLE_REENTER_ADD_ROWS_TO_SHEET,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createSheetOperationTags(
                (List<String>) map.get(ROWS), (List<File>) map.get(ROWS_FROM_CSV_FILE));
        telemetryHelper.incrementCount("office365Excel.sheetReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            List<String> rows = (List<String>) map.get(ROWS);
            List<File> rowsFromFile = (List<File>) map.get(ROWS_FROM_CSV_FILE);

            LOGGER.info("Processing corrected input for driveId: {}, fileId: {}, sheetName: {}", driveId, fileId, sheetName);
            List<FreeForm> result = this.worksheetService.addRows(driveId, fileId, sheetName, rows, rowsFromFile, null);

            LOGGER.info("Successfully added rows to sheet: {}", sheetName);
            return ExtensionResponseFactory.create(Map.of("Sheet Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.sheetReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to add rows on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding rows to the sheet. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.sheetReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteRowWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleDeleteRowWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter delete row data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.deleteRowWithIndexConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_DELETE_ROW_WITH_INDEX,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteRowWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleDeleteRowWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteRowWithIndexOperationTags(
                (String) map.get(SHEET_NAME), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.deleteRowWithIndexReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            Double rowIndex = (Double) map.get(ROW_INDEX);

            LOGGER.info("Processing corrected input for delete row with index operation");

            Response deleteResponse = restClientProvider.getRestClient(driveId).post(
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
    public ExtensionResponse handleGetCompleteRowInfoConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleGetCompleteRowInfoConfirmation start: {}", map);
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

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.getCompleteRowInfoConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_COMPLETE_ROW_INFO,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetCompleteRowInfoReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleGetCompleteRowInfoReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetCompleteRowInfoOperationTags(
                (String) map.get(SHEET_NAME), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.getCompleteRowInfoReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            Double rowIndex = (Double) map.get(ROW_INDEX);

            LOGGER.info("Processing corrected input for get complete row info operation");

            FreeForm result = this.worksheetService.getRow(fileId, sheetName, rowIndex, driveId);

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
    public ExtensionResponse handleSearchConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleSearchConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter search data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.searchConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_SEARCH,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleSearchReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleSearchReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createSearchOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(DATA_TO_SEARCH));
        telemetryHelper.incrementCount("office365Excel.searchReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String dataToSearch = (String) map.get(DATA_TO_SEARCH);

            LOGGER.info("Processing corrected input for search operation");
            List<String> result = this.worksheetService.findValue(driveId, fileId, sheetName, dataToSearch);

            LOGGER.info("Successfully searched for value: {}", dataToSearch);
            return ExtensionResponseFactory.create(Map.of("Cell Ids", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.searchReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to search on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while searching. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.searchReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteColumnWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleDeleteColumnWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter delete column data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.deleteColumnWithIndexConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_DELETE_COLUMN_WITH_INDEX,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteColumnWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleDeleteColumnWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteColumnOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(COLUMN_NAME));
        telemetryHelper.incrementCount("office365Excel.deleteColumnWithIndexReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String columnName = (String) map.get(COLUMN_NAME);

            LOGGER.info("Processing corrected input for delete column with index operation");

            Response deleteResponse = restClientProvider.getRestClient(driveId).post(
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

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleAddRowWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter add row with index data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults, null, Map.of());
        }

        LOGGER.info("User chose to re-enter add row with index data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.addRowWithIndexConfirmation");
        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(DRIVE_ID, state.get(DRIVE_ID));
        stateMap.put(FILE_ID, state.get(FILE_ID));
        stateMap.put(SHEET_NAME, state.get(SHEET_NAME));
        stateMap.put(ROWS, state.get(ROWS));
        stateMap.put(ROW_INDEX, state.get(ROW_INDEX));

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
                SubCatalogConstants.HANDLE_REENTER_ADD_ROW_WITH_INDEX,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleAddRowWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddRowWithIndexOperationTags(
                (String) map.get(SHEET_NAME), (List<String>) map.get(ROWS), (List<File>) map.get(ROWS_FROM_CSV_FILE), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.addRowWithIndexReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            List<String> rows = (List<String>) map.get(ROWS);
            List<File> rowsFromFile = (List<File>) map.get(ROWS_FROM_CSV_FILE);
            Double rowIndex = (Double) map.get(ROW_INDEX);

            LOGGER.info("Processing corrected input for add row with index operation");

            List<String> rowsStringOrCsvRows = stringUtil.getStringList(rows, rowsFromFile);

            restClientProvider.getRestClient(driveId).post(fileId +
                            Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                            + stringUtil.calculateAddressForInsert(rowIndex, rows,
                            rowsFromFile) + "')/insert",
                    Map.of(Constants.SHIFT, Constants.DOWN));
            sheetFormattingService.clearRowFormatting(driveId, fileId, sheetName, rowIndex);
            StringUtil.appendCommas(rowsStringOrCsvRows);

            List<FreeForm> result = worksheetService.addRows(driveId, fileId, sheetName, rowsStringOrCsvRows,
                    StringUtil.calculateAddressForInsert(rowIndex, rowsStringOrCsvRows));

            LOGGER.info("Successfully added row at index: {}", rowIndex);
            return ExtensionResponseFactory.create(Map.of("Sheet Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addRowWithIndexReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to add row on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding row at the specified index. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addRowWithIndexReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleUpdateRowWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleUpdateRowWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter update row with index data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults, null, Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.updateRowWithIndexConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_UPDATE_ROW_WITH_INDEX,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleUpdateRowWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleUpdateRowWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createUpdateRowWithIndexOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(ROW), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.updateRowWithIndexReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String row = (String) map.get(ROW);
            Double rowIndex = (Double) map.get(ROW_INDEX);

            if (row == null || row.isBlank()) {
                LOGGER.debug("Row data not provided in reentry for driveId: {}, fileId: {}, sheetName: {}", driveId, fileId, sheetName);
                throw new IllegalArgumentException("Row data not provided");
            }

            LOGGER.info("Processing corrected input for driveId: {}, fileId: {}, sheetName: {}, rowIndex: {}",
                    driveId, fileId, sheetName, rowIndex);

            List<FreeForm> result = worksheetService.addRows(driveId, fileId, sheetName,
                    List.of(row), StringUtil.calculateAddressForInsert(rowIndex, List.of(row)));

            LOGGER.info("Successfully updated row at index: {}", rowIndex);
            return ExtensionResponseFactory.create(Map.of("Sheet Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.updateRowWithIndexReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
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
    public ExtensionResponse handleGetColumnValuesInSheetConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleGetColumnValuesInSheetConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter column values in sheet data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.getColumnValuesInSheetConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES_IN_SHEET,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetColumnValuesInSheetReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleGetColumnValuesInSheetReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();

        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetColumnValuesForKeysInSheetOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(RANGE), (List<String>) map.get(LOOKUP_ROWS), (List<String>) map.get(LOOKUP_COLUMNS),
                (Boolean) map.get(GET_ALL_RECORD), (Double) map.get(COLUMN_NUMBER));
        telemetryHelper.incrementCount("office365Excel.getColumnValuesInSheetReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String range = (String) map.get(RANGE);
            List<String> lookupRows = (List<String>) map.get(LOOKUP_ROWS);
            List<String> lookupColumns = (List<String>) map.get(LOOKUP_COLUMNS);
            Boolean getAllRecord = (Boolean) map.get(GET_ALL_RECORD);
            Double columnNumber = (Double) map.get(COLUMN_NUMBER);

            LOGGER.info("Processing corrected input for get column values in sheet operation");
            List<FreeForm> result = this.worksheetOperation.getRowAndColumnData(fileId, null, lookupRows, lookupColumns,
                    getAllRecord, columnNumber, sheetName, range, driveId);

            LOGGER.info("Successfully retrieved column values for keys in sheet: {}", safeString(sheetName));
            return ExtensionResponseFactory.create(Map.of("Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getColumnValuesInSheetReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get column values on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving column values. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getColumnValuesInSheetReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleClearSharedSheetConfirmation(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleClearSharedSheetConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter clear shared sheet data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.clearSharedSheetConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_CLEAR_SHARED_SHEET,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleClearSharedSheetReentry(Map<String, Object> map) {
        LOGGER.info("SheetReentryHandler.handleClearSharedSheetReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createClearSharedSheetForRangeOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(RANGE));
        telemetryHelper.incrementCount("office365Excel.clearSharedSheetReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String cellRange = (String) map.get(RANGE);

            LOGGER.info("Processing corrected input for clear shared sheet operation");
            Boolean result = worksheetService.clearWorksheetCells(driveId, fileId, sheetName, cellRange);

            LOGGER.info("Successfully cleared sheet range on retry");
            return ExtensionResponseFactory.create(Map.of("Cleared", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.clearSharedSheetReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to clear sheet range on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while clearing the sheet range. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.clearSharedSheetReentry", startTime, telemetryTags);
        }
    }
}
