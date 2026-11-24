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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelSheet;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.ExtensionResponseFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.RemediationActionFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.WorksheetReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.SheetFormattingService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.*;
import java.util.stream.Collectors;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

/**
 * Implementation for worksheet-specific reentry operations
 */
public class WorksheetReentryHandlerImpl extends BaseReentryHandlerImpl implements WorksheetReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(WorksheetReentryHandlerImpl.class);
    private final WorksheetService worksheetService;
    private final RestClientProvider restClientProvider;
    private final StringUtil stringUtil;
    private final SheetFormattingService sheetFormattingService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public WorksheetReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                       WorksheetService worksheetService, RestClientProvider restClientProvider,
                                       StringUtil stringUtil, SheetFormattingService sheetFormattingService,
                                       TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.worksheetService = worksheetService;
        this.restClientProvider = restClientProvider;
        this.stringUtil = stringUtil;
        this.sheetFormattingService = sheetFormattingService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteSheetConfirmation(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleDeleteSheetConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter delete sheet data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter delete sheet data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.deleteSheetConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_DELETE_SHEET,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleDeleteSheetReentry(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleDeleteSheetReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteSheetOperationTags((String) map.get(SHEET_NAME));
        telemetryHelper.incrementCount("office365Excel.deleteSheetReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);

            LOGGER.info("Processing corrected input for delete sheet operation");
            String result = worksheetService.deleteWorksheets(fileId, sheetName);

            LOGGER.info("Successfully deleted sheet: {}", sheetName);
            return ExtensionResponseFactory.create(Map.of("Result", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.deleteSheetReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to delete sheet on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while deleting the sheet. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.deleteSheetReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowsConfirmation(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleAddRowsConfirmation start: {}", map);
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

        LOGGER.info("User chose to re-enter add rows data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.addRowsConfirmation");
        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(FILE_ID, state.get(FILE_ID));
        stateMap.put(SHEET_NAME, state.get(SHEET_NAME));
        stateMap.put(ROWS, state.get(ROWS));

        Object fileObj = map.get(ROWS_FROM_CSV_FILE);
        List<File> rowsFromFileList = extractFilesFromMap(fileObj);

        stateMap.put(
                ROWS_FROM_CSV_FILE,
                !rowsFromFileList.isEmpty() ? rowsFromFileList : state.get(ROWS_FROM_CSV_FILE)
        );

        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_ADD_ROWS,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowsReentry(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleAddRowsReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createSheetOperationTags(
                (List<String>) map.get(ROWS), (List<File>) map.get(ROWS_FROM_CSV_FILE));
        telemetryHelper.incrementCount("office365Excel.addRowsReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            List<String> rows = (List<String>) map.get(ROWS);
            List<File> rowsFromFile = (List<File>) map.get(ROWS_FROM_CSV_FILE);

            if (!Worksheets.isRowDataAvailable(rows, rowsFromFile)) {
                LOGGER.debug("Row data not provided in reentry for fileId: {}, sheetName: {}", fileId, sheetName);
                throw new IllegalArgumentException("Row data not provided");
            }

            LOGGER.info("Processing corrected input for fileId: {}, sheetName: {}", fileId, sheetName);
            List<FreeForm> result = this.worksheetService.addRows(null, fileId, sheetName, rows, rowsFromFile, null);

            LOGGER.info("Successfully added rows to sheet: {}", sheetName);
            return ExtensionResponseFactory.create(Map.of("Sheet Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addRowsReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add rows on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding rows to the sheet. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addRowsReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowWithIndexConfirmation(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleAddRowWithIndexConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter add row with index data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter add row with index data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.addRowWithIndexConfirmation");
        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(FILE_ID, state.get(FILE_ID));
        stateMap.put(SHEET_NAME, state.get(SHEET_NAME));
        stateMap.put(ROW_INDEX, state.get(ROW_INDEX));
        stateMap.put(ROWS, state.get(ROWS));
        Object fileObj = map.get(ROWS_FROM_CSV_FILE);
        List<File> rowsFromFileList = extractFilesFromMap(fileObj);

        stateMap.put(
                ROWS_FROM_CSV_FILE,
                !rowsFromFileList.isEmpty() ? rowsFromFileList : state.get(ROWS_FROM_CSV_FILE)
        );

        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_ADD_SHEET_ROW_WITH_INDEX,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddRowWithIndexReentry(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleAddRowWithIndexReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddRowWithIndexOperationTags(
                (String) map.get(SHEET_NAME), (List<String>) map.get(ROWS), (List<File>) map.get(ROWS_FROM_CSV_FILE), (Double) map.get(ROW_INDEX));
        telemetryHelper.incrementCount("office365Excel.addRowWithIndexReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            Double rowIndex = (Double) map.get(ROW_INDEX);
            List<String> rows = (List<String>) map.get(ROWS);
            List<File> rowsFromFile = (List<File>) map.get(ROWS_FROM_CSV_FILE);

            LOGGER.info("Processing corrected input for add row with index operation");
            List<String> rowsStringOrCsvRows = stringUtil.getStringList(rows, rowsFromFile);

            restClientProvider.getRestClient(null).post(fileId +
                            Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                            + stringUtil.calculateAddressForInsert(rowIndex, rows, rowsFromFile) + "')/insert",
                    Map.of(Constants.SHIFT, Constants.DOWN));

            sheetFormattingService.clearRowFormatting(fileId, sheetName, rowIndex, rowsStringOrCsvRows.size());
            StringUtil.appendCommas(rowsStringOrCsvRows);

            List<FreeForm> result = worksheetService.addRows(null, fileId, sheetName, rowsStringOrCsvRows,
                    StringUtil.calculateAddressForInsert(rowIndex, rowsStringOrCsvRows));

            LOGGER.info("Successfully added row at index: {}", rowIndex);
            return ExtensionResponseFactory.create(Map.of("Sheet Data", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addRowWithIndexReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add row with index on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding the row. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addRowWithIndexReentry", startTime, telemetryTags);
        }
    }

    private List<File> extractFilesFromMap(Object fileObj) {
        if (fileObj instanceof List<?> list) {
            List<File> files = list.stream()
                    .filter(File.class::isInstance)
                    .map(File.class::cast)
                    .collect(Collectors.toList());
            LOGGER.debug("Extracted {} file(s) from list", files.size());
            return files;
        } else if (fileObj instanceof File file) {
            LOGGER.debug("Extracted single file: {}", file.getFileName());
            return List.of(file);
        } else {
            LOGGER.debug("No valid files found in the input map");
            return Collections.emptyList();
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetSheetsInfoConfirmation(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleGetSheetsInfoConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter sheets information. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter sheets information. Returning fetch prompt response.");
        telemetryHelper.incrementCount("office365Excel.getSheetsInfoConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_GET_SHEETS_INFO,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetSheetsInfoReentry(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleGetSheetsInfoReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetSheetsInfoOperationTags((String) map.get(FILE_ID));
        telemetryHelper.incrementCount("office365Excel.getSheetsInfoReentry");

        try {
            String fileId = (String) map.get(FILE_ID);

            LOGGER.info("Processing corrected input for fileId: {}", fileId);
            List<ExcelSheet> result = this.worksheetService.getWorksheets(fileId, null, null, null);

            LOGGER.info("Successfully retrieved sheets information for file: {}", fileId);
            return ExtensionResponseFactory.create(Map.of("Sheets", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getSheetsInfoReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get sheets information on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving sheets information. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getSheetsInfoReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleCreateSheetConfirmation(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleCreateSheetConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter create sheet data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter create sheet data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.createSheetConfirmation");
        Map<String, Object> stateMap = new HashMap<>();
        stateMap.put(STATE_ID, stateId);
        stateMap.put(FILE_ID, state.get(FILE_ID));
        stateMap.put(SHEET_NAME, state.get(SHEET_NAME));

        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                HANDLE_REENTER_CREATE_SHEET,
                stateMap);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleCreateSheetReentry(Map<String, Object> map) {
        LOGGER.info("WorksheetReentryHandler.handleCreateSheetReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createCreateSheetOperationTags((String) map.get(SHEET_NAME));
        telemetryHelper.incrementCount("office365Excel.createSheetReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);

            LOGGER.info("Processing corrected input for create sheet operation");
            ExcelSheet result = worksheetService.createWorksheet(fileId, sheetName);

            LOGGER.info("Successfully created sheet: {}", sheetName);
            return ExtensionResponseFactory.create(Map.of("Sheet", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.createSheetReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to create sheet on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while creating the sheet. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.createSheetReentry", startTime, telemetryTags);
        }
    }

}