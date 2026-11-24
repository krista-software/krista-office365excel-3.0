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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.FormulaReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.model.base.FreeForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.REENTER;

/**
 * Implementation for formula-related reentry operations
 */
public class FormulaReentryHandlerImpl extends BaseReentryHandlerImpl implements FormulaReentryHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(FormulaReentryHandlerImpl.class);

    private final WorkbookService workbookService;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public FormulaReentryHandlerImpl(WorkbookService workbookService,
                                     ErrorHandlingStateManager internalStateManager,
                                     ExtensionResponseGenerator responseGenerator,
                                     TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.workbookService = workbookService;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleExecuteFormulaRetryConfirmation(Map<String, Object> map) {
        LOGGER.info("FormulaReentryHandler.handleExecuteFormulaRetryConfirmation start: {}", map);

        Boolean retry = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(retry)) {
            LOGGER.info("User chose not to retry execute formula. Returning fetch deny response.");
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("Retry was true, hence continuing");
        telemetryHelper.incrementCount("office365Excel.executeFormulaRetryConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_RETRY_EXECUTE_FORMULA,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleExecuteFormulaRetry(Map<String, Object> map) {
        LOGGER.info("FormulaReentryHandler.handleExecuteFormulaRetry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createExecuteFormulaOperationTags(
                (String) map.get(FORMULA), (String) map.get(FORMULA_NAME));
        telemetryHelper.incrementCount("office365Excel.executeFormulaRetry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String formula = (String) map.get(FORMULA);
            String formulaName = (String) map.get(FORMULA_NAME);
            String comment = (String) map.get(COMMENT);

            LOGGER.info("Processing retry for execute formula operation");
            FreeForm result = this.workbookService.executeFormula(null, fileId, formula, formulaName, comment);

            LOGGER.info("Successfully executed formula after retry: {}", safeString(formula));
            return ExtensionResponseFactory.create(Map.of("Formula Result", result));
        } catch (Exception cause) {
            LOGGER.error("Retry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.executeFormulaRetry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to execute formula on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while executing the formula. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.executeFormulaRetry", startTime, telemetryTags);
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddFormulaToCellConfirmation(Map<String, Object> map) {
        LOGGER.info("FormulaReentryHandler.handleAddFormulaToCellConfirmation start: {}", map);

        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter add formula to cell data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info("User chose to re-enter add formula to cell data. Preparing fetch response.");
        telemetryHelper.incrementCount("office365Excel.addFormulaToCellConfirmation");
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_ADD_FORMULA_TO_CELL,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleAddFormulaToCellReentry(Map<String, Object> map) {
        LOGGER.info("FormulaReentryHandler.handleAddFormulaToCellReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddFormulaToCellOperationTags(
                (String) map.get(CELL_ID), (String) map.get(FORMULA));
        telemetryHelper.incrementCount("office365Excel.addFormulaToCellReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String sheetName = (String) map.get(SHEET_NAME);
            String cellId = (String) map.get(CELL_ID);
            String formula = (String) map.get(FORMULA);

            LOGGER.info("Processing corrected input for add formula to cell operation");
            FreeForm result = this.workbookService.addFormula(fileId, sheetName, cellId, formula);

            LOGGER.info("Successfully added formula to cell: {}", safeString(cellId));
            return ExtensionResponseFactory.create(Map.of("Cell Information", result));
        } catch (Exception cause) {
            LOGGER.error("Error processing add formula to cell reentry: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.addFormulaToCellReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to add formula to cell: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while executing the formula. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.addFormulaToCellReentry", startTime, telemetryTags);
        }
    }
}