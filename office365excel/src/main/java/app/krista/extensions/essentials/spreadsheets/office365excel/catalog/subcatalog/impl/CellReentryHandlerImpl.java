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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.CellReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;

import javax.inject.Inject;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER_WAS_TRUE_HENCE_CONTINUING;

/**
 * Implementation for cell-specific reentry operations
 */
public class CellReentryHandlerImpl extends BaseReentryHandlerImpl implements CellReentryHandler {

    private final WorksheetService worksheetService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public CellReentryHandlerImpl(ErrorHandlingStateManager internalStateManager,
                                  ExtensionResponseGenerator responseGenerator,
                                  WorksheetService worksheetService,
                                  TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.worksheetService = worksheetService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddDataToCellConfirmation(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddDataToCellConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter cell data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.addDataToCellConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_ADD_DATA_TO_CELL,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddDataToCellReentry(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddDataToCellReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        // Telemetry tracking for operation initiation
        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createCellOperationTags((String) map.get(CELL_ID), (String) map.get(DATA));
        telemetryHelper.incrementCount("office365Excel.addDataToCellReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String cellId = (String) map.get(CELL_ID);
            String data = (String) map.get(DATA);

            LOGGER.info("Processing corrected input for add data to cell operation");
            FreeForm result = this.worksheetService.addDataToCell(driveId, fileId, sheetName, cellId, data);

            LOGGER.info("Successfully added data to cell: {}", cellId);
            return ExtensionResponseFactory.create(Map.of("Cell Information", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addDataToCellReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add data to cell on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding data to cell. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addDataToCellReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleCellDataConfirmation(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleCellDataConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter cell data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.cellDataConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_CELL_DATA,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleCellDataReentry(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleCellDataReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createCellDataOperationTags((String) map.get(SHEET_NAME), (String) map.get(CELL_ID));
        telemetryHelper.incrementCount("office365Excel.cellDataReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String cellId = (String) map.get(CELL_ID);

            LOGGER.info("Processing corrected input for cell data operation");
            FreeForm result = this.worksheetService.getCellData(driveId, fileId, sheetName, cellId);

            LOGGER.info("Successfully retrieved data for cell: {}", cellId);
            return ExtensionResponseFactory.create(Map.of("Cell Information", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.cellDataReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get cell data on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving cell data. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.cellDataReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddMultipleCellDataConfirmation(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddMultipleCellDataConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter multiple cell data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.addMultipleCellDataConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddMultipleCellDataReentry(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddMultipleCellDataReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        List<Map<String, Object>> cellIDData = (List<Map<String, Object>>) map.get(CELL_ID_DATA);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddMultipleCellDataOperationTags(
                (String) map.get(SHEET_NAME), cellIDData);
        telemetryHelper.incrementCount("office365Excel.addMultipleCellDataReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);

            LOGGER.info("Processing corrected input for add multiple cell data operation");
            List<FreeForm> allCellData = new ArrayList<>();

            for (Map eachMap : cellIDData) {
                String cellId = (String) eachMap.get(CELL_ID);
                if (!Worksheets.isValidCellId(cellId)) {
                    throw new IllegalArgumentException(Constants.INVALID_CELL_ID_IS_GIVEN + cellId);
                }
                String dataToBeAdded = (String) eachMap.get(Constants.DATA);
                allCellData.add(this.worksheetService.addDataToCell(driveId, fileId, sheetName, cellId, dataToBeAdded));
            }

            LOGGER.info("Successfully added data to multiple cells on retry");
            return ExtensionResponseFactory.create(Map.of("Cell Data", allCellData));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addMultipleCellDataReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add multiple cell data on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding data to multiple cells. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addMultipleCellDataReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddMultipleCellDataConfirmationForSheet(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddMultipleCellDataConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter multiple cell data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.addMultipleCellDataConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddMultipleCellDataReentryForSheet(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddMultipleCellDataReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        List<Map<String, Object>> cellIDData = (List<Map<String, Object>>) map.get(CELL_ID_DATA);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddMultipleCellDataOperationTags(
                (String) map.get(SHEET_NAME), cellIDData);
        telemetryHelper.incrementCount("office365Excel.addMultipleCellDataReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);

            LOGGER.info("Processing corrected input for add multiple cell data operation");
            List<FreeForm> allCellData = new ArrayList<>();

            for (Map eachMap : cellIDData) {
                String cellId = (String) eachMap.get(CELL_ID);
                if (!Worksheets.isValidCellId(cellId)) {
                    throw new IllegalArgumentException(Constants.INVALID_CELL_ID_IS_GIVEN + cellId);
                }
                String dataToBeAdded = (String) eachMap.get(Constants.DATA);
                allCellData.add(this.worksheetService.addDataToCell(null, fileId, sheetName, cellId, dataToBeAdded));
            }

            LOGGER.info("Successfully added data to multiple cells on retry");
            return ExtensionResponseFactory.create(Map.of("Cell Data", allCellData));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addMultipleCellDataReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add multiple cell data on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding data to multiple cells. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addMultipleCellDataReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetCellValueConfirmation(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleGetCellValueConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter get cell value data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter get cell value data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.getCellValueConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_CELL_VALUE,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetCellValueReentry(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleGetCellValueReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetCellValueOperationTags(
                (String) map.get(FILE_ID), (String) map.get(SHEET_NAME), (String) map.get(CELL_ID));
        telemetryHelper.incrementCount("office365Excel.getCellValueReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String cellId = (String) map.get(CELL_ID);

            LOGGER.info("Processing corrected input for get cell value operation");
            FreeForm result = this.worksheetService.getCellData(null, fileId, sheetName, cellId);
            LOGGER.info("Successfully retrieved cell value after reentry");
            return ExtensionResponseFactory.create(Map.of("Cell Information", result));
        } catch (Exception cause) {
            LOGGER.error("Error in handleGetCellValueReentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getCellValueReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause, "Failed to get cell value: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving the cell value. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getCellValueReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddDataToCellRetryConfirmation(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddDataToCellRetryConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to retry add data to cell. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to retry add data to cell. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.addDataToCellRetryConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_RETRY_ADD_DATA_TO_CELL,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddDataToCellRetry(Map<String, Object> map) {
        LOGGER.info("CellReentryHandler.handleAddDataToCellRetry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddDataToCellOperationTags(
                (String) map.get(CELL_ID), (String) map.get(DATA));
        telemetryHelper.incrementCount("office365Excel.addDataToCellRetry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String cellId = (String) map.get(CELL_ID);
            String data = (String) map.get(DATA);

            LOGGER.info("Processing retry for add data to cell operation");
            FreeForm result = this.worksheetService.addDataToCell(null, fileId, sheetName, cellId, data);

            LOGGER.info("Successfully added data to cell after retry: {}", cellId);
            return ExtensionResponseFactory.create(Map.of("Cell Information", result));
        } catch (Exception cause) {
            LOGGER.error("Retry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addDataToCellRetry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add data to cell on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while adding data to cell. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addDataToCellRetry", startTime, telemetryTags);
        }
    }
}
