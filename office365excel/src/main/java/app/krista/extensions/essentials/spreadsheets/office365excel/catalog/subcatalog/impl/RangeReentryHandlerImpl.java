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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.RangeReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.model.base.FreeForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

/**
 * Implementation for range-specific reentry operations
 */
public class RangeReentryHandlerImpl extends BaseReentryHandlerImpl implements RangeReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(RangeReentryHandlerImpl.class);
    private final WorkbookService workbookService;
    private final WorksheetService worksheetService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public RangeReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                   WorkbookService workbookService, WorksheetService worksheetService, TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.workbookService = workbookService;
        this.worksheetService = worksheetService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleConfirmation(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter range sum data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter range sum data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.rangeSumConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_RANGE_SUM,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReentry(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(RANGE));
        telemetryHelper.incrementCount("office365Excel.rangeReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String range = (String) map.get(RANGE);

            LOGGER.info("Processing corrected input for driveId: {}, fileId: {}, sheetName: {}, range: {}",
                    driveId, fileId, sheetName, range);
            Double result = this.workbookService.worksheetRangeSum(driveId, fileId, sheetName, range);

            LOGGER.info("Successfully calculated sum for range: {}", range);
            return ExtensionResponseFactory.create(Map.of("Sum", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.rangeReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to calculate range sum on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while calculating the sum. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.rangeReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleSheetVlookupConfirmation(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleSheetVlookupConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter vlookup data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter vlookup data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.sheetVlookupConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_SHARED_SHEET_VLOOKUP,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleSheetVlookupReentry(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleSheetVlookupReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createVlookupOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(LOOKUP_VALUE), (String) map.get(RANGE));
        telemetryHelper.incrementCount("office365Excel.sheetVlookupReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String lookupValue = (String) map.get(LOOKUP_VALUE);
            String sheetName = (String) map.get(SHEET_NAME);
            String range = (String) map.get(RANGE);
            Double columnIndex = (Double) map.get(COLUMN_INDEX);
            Boolean exactMatch = (Boolean) map.get(EXACT_MATCH);

            LOGGER.info("Processing corrected input for vlookup operation");
            Object result = this.workbookService.vLookup(driveId, fileId, lookupValue, sheetName, range, columnIndex, exactMatch);

            LOGGER.info("Successfully performed VLOOKUP for value: {} in range: {}", lookupValue, range);
            return ExtensionResponseFactory.create(Map.of("Result", result));
        } catch (Exception cause) {
            LOGGER.error("VLOOKUP reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.sheetVlookupReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(cause,
                    "Failed to perform VLOOKUP on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while performing VLOOKUP. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.sheetVlookupReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetRowsForRangeConfirmation(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleGetRowsForRangeConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter range data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        telemetryHelper.incrementCount("office365Excel.getRowsForRangeConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleGetRowsForRangeReentry(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleGetRowsForRangeReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags(
                (String) map.get(SHEET_NAME), (String) map.get(RANGE));
        telemetryHelper.incrementCount("office365Excel.getRowsForRangeReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String range = (String) map.get(RANGE);

            LOGGER.info("Processing corrected input for get rows for range operation");
            List<FreeForm> result = this.worksheetService.getRows(driveId, fileId, sheetName, range);

            LOGGER.info("Successfully retrieved rows for range: {}", range);
            return ExtensionResponseFactory.create(Map.of("Rows", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.getRowsForRangeReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to get rows for range on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while retrieving rows. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.getRowsForRangeReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleExecuteFormulaConfirmation(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleExecuteFormulaConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter formula data. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter formula data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.executeFormulaConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_EXECUTE_FORMULA,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleExecuteFormulaReentry(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleExecuteFormulaReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createExecuteFormulaOperationTags(
                (String) map.get(FORMULA), (String) map.get(FORMULA_NAME));
        telemetryHelper.incrementCount("office365Excel.executeFormulaReentry");

        try {
            String driveId = (String) map.get(DRIVE_ID);
            String fileId = (String) map.get(FILE_ID);
            String formula = (String) map.get(FORMULA);
            String formulaName = (String) map.get(FORMULA_NAME);
            String comment = (String) map.get(COMMENT);

            LOGGER.info("Processing corrected input for execute formula operation");
            FreeForm result = this.workbookService.executeFormula(driveId, fileId, formula, formulaName, comment);

            LOGGER.info("Successfully executed formula: {}", formula);
            return ExtensionResponseFactory.create(Map.of("Formula Result", result));
        } catch (Exception cause) {
            LOGGER.error("Reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.executeFormulaReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to execute formula on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while executing the formula. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.executeFormulaReentry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleClearSharedSheetConfirmation(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleClearSharedSheetConfirmation start: {}", map);
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

        LOGGER.info("Reenter was true, hence continuing");
        telemetryHelper.incrementCount("office365Excel.clearSharedSheetConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_CLEAR_SHEET,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleClearSharedSheetReentry(Map<String, Object> map) {
        LOGGER.info("RangeReentryHandler.handleClearSharedSheetReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createClearSheetForRangeOperationTagsForSheet(
                (String) map.get(SHEET_NAME), (String) map.get(RANGE));
        telemetryHelper.incrementCount("office365Excel.clearSharedSheetReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String cellRange = (String) map.get(RANGE);

            LOGGER.info("Processing corrected input for clear shared sheet operation");
            Boolean result = worksheetService.clearWorksheetCells(fileId, sheetName, cellRange);

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
