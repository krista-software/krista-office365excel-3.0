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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.FreeForm;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class FormulaResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(FormulaResponseService.class);

    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    private final WorkbookService workbookService;

    @Inject
    public FormulaResponseService(ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager,
                                  ExtensionResponseGenerator responseGenerator, RequestContext requestContext,
                                  TelemetryHelper telemetryHelper, WorkbookService workbookService) {
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
        this.workbookService = workbookService;
    }

    public ExtensionResponse getExecuteFormulaExtensionResponse(String fileId, String formula, String formulaName, String comment, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.executeFormula");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createExecuteFormulaValidationMapForSheet(fileId, formula, formulaName, comment);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createExecuteFormulaOperationTags(formula, formulaName);

        return handler.execute(
                "executeFormula",
                validationMap,
                Map.of(FILE_ID, fileId, FORMULA, formula, FORMULA_NAME, formulaName, COMMENT, comment),
                SubCatalogConstants.CONFIRM_RETRY_EXECUTE_FORMULA,
                stateId -> StateMapperUtil.executeFormulaMetaToMap(null, fileId, formula, formulaName, comment, stateId),
                () -> this.workbookService.executeFormula(null, fileId, formula, formulaName, comment),
                result -> {
                    LOGGER.info("Successfully executed formula: {}", safeString(formula));
                    return Map.of("Formula Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getVlookupExtensionResponse(String fileID, String lookupValue, String sheetName, String range, Double columnIndex, Boolean exactMatch, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.vlookup");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createVlookupValidationMapForSheet(fileID, sheetName, range, columnIndex, lookupValue);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createVlookupOperationTags(lookupValue, range, sheetName);

        return handler.execute(
                "vlookup",
                validationMap,
                Map.of(FILE_ID, fileID, LOOKUP_VALUE, lookupValue, SHEET_NAME, sheetName, RANGE, range,
                        COLUMN_INDEX, columnIndex, EXACT_MATCH, exactMatch != null ? exactMatch : false),
                SubCatalogConstants.CONFIRM_REENTER_SHEET_VLOOKUP,
                stateId -> StateMapperUtil.vlookupMetaToMap(fileID, lookupValue, sheetName, range, columnIndex, exactMatch, stateId),
                () -> workbookService.vLookup(null, fileID, lookupValue, sheetName, range, columnIndex, exactMatch),
                result -> {
                    LOGGER.info("Successfully performed vlookup for value: {} in range: {}", safeString(lookupValue), safeString(range));
                    return Map.of("Lookup Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getSummationExtensionResponse(String fileID, String sheetName, String range, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.summation");

        AreaOperationHandler<Double> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createRangeValidationMapForSheet(fileID, sheetName, range);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createRangeOperationTags(range, sheetName);

        return handler.execute(
                "summation",
                validationMap,
                Map.of(FILE_ID, fileID, SHEET_NAME, sheetName, RANGE, range),
                SubCatalogConstants.CONFIRM_REENTER_SUMMATION,
                stateId -> StateMapperUtil.summationMetaToMap(fileID, sheetName, range, stateId),
                () -> this.workbookService.worksheetRangeSum(null, fileID, sheetName, range),
                result -> {
                    LOGGER.info("Successfully calculated sum for range: {}", safeString(range));
                    return Map.of("Sum", result);
                },
                telemetryTags,
                allowRetry
        );
    }

}
