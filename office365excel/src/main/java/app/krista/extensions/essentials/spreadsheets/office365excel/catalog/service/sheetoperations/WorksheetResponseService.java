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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelSheet;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.ksdk.context.RequestContext;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class WorksheetResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(WorksheetResponseService.class);

    private final WorksheetService worksheetService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public WorksheetResponseService(WorksheetService worksheetService, ValidationOrchestrator validationOrchestrator,
                                    ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                    RequestContext requestContext, TelemetryHelper telemetryHelper) {
        this.worksheetService = worksheetService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    public ExtensionResponse getDeleteSheetExtensionResponse(String fileId, String sheetName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.deleteSheet");

        AreaOperationHandler<String> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createDeleteSheetValidationMap(fileId, sheetName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createDeleteSheetOperationTags(sheetName);

        return handler.execute(
                "deleteSheet",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName),
                SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET,
                stateId -> StateMapperUtil.deleteSheetMetaToMap(fileId, sheetName, stateId),
                () -> this.worksheetService.deleteWorksheets(fileId, sheetName),
                result -> {
                    LOGGER.info("Successfully deleted sheet: {}", safeString(sheetName));
                    return Map.of("Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getCreateSheetExtensionResponse(String fileId, String newSheetName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.createSheet");

        AreaOperationHandler<ExcelSheet> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createCreateSheetValidationMap(fileId);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createCreateSheetOperationTags(newSheetName);

        return handler.execute(
                "createSheet",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, newSheetName),
                SubCatalogConstants.CONFIRM_REENTER_CREATE_SHEET,
                stateId -> StateMapperUtil.createSheetMetaToMap(fileId, newSheetName, stateId),
                () -> worksheetService.createWorksheet(fileId, newSheetName),
                result -> {
                    LOGGER.info("Successfully created sheet: {}", safeString(newSheetName));
                    return Map.of("Sheet", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getGetSheetsInformationExtensionResponse(String fileId, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getSheetsInformation");

        AreaOperationHandler<List<ExcelSheet>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetSheetsInfoValidationMap(fileId);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetSheetsInfoOperationTags(fileId);

        return handler.execute(
                "getSheetsInformation",
                validationMap,
                Map.of(FILE_ID, fileId),
                SubCatalogConstants.CONFIRM_REENTER_GET_SHEETS_INFO,
                stateId -> StateMapperUtil.getSheetsInfoMetaToMap(fileId, stateId),
                () -> this.worksheetService.getWorksheets(fileId, null, null, null),
                result -> {
                    LOGGER.info("Successfully retrieved sheets information for file: {}", safeString(fileId));
                    return Map.of("Sheets", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getClearSheetForRangeExtensionResponse(String fileID, String sheetName, String cellRange, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.clearSheetForRange");

        AreaOperationHandler<Boolean> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createClearSheetForRangeValidationMap(fileID, sheetName, cellRange);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createClearSheetForRangeOperationTagsForSheet(sheetName, cellRange);

        return handler.execute(
                "clearSheetForRange",
                validationMap,
                Map.of(FILE_ID, fileID, SHEET_NAME, sheetName, RANGE, cellRange),
                SubCatalogConstants.CONFIRM_REENTER_CLEAR_SHEET,
                stateId -> StateMapperUtil.clearSheetForRangeMetaToMap(fileID, sheetName, cellRange, stateId),
                () -> worksheetService.clearWorksheetCells(fileID, sheetName, cellRange),
                result -> {
                    LOGGER.info("Successfully cleared sheet range: {}", safeString(cellRange));
                    return Map.of("Cleared", result);
                },
                telemetryTags,
                allowRetry
        );
    }
}
