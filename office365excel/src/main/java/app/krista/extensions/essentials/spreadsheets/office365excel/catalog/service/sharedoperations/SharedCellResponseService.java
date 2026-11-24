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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.FreeForm;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

/**
 * Helper class for creating extension responses for shared cell operations
 */
@Service
public class SharedCellResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(SharedCellResponseService.class);

    private final WorksheetService worksheetService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;


    @Inject
    public SharedCellResponseService(WorksheetService worksheetService, ValidationOrchestrator validationOrchestrator,
                                     ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                     RequestContext requestContext, TelemetryHelper telemetryHelper) {
        this.worksheetService = worksheetService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    public ExtensionResponse getAddDataToCellExtensionResponse(String driveId, String fileId, String sheetName, String cellId, String data, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addDataToProvidedCell");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createCellValidationMap(driveId, fileId, sheetName, cellId);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createCellOperationTags(cellId, data);

        return handler.execute(
                "addDataToProvidedCell",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, CELL_ID, cellId, DATA, data),
                SubCatalogConstants.CONFIRM_REENTER_ADD_DATA_TO_CELL,
                stateId -> StateMapperUtil.addDataToCellMetaToMap(driveId, fileId, sheetName, cellId, data, stateId),
                () -> worksheetService.addDataToCell(driveId, fileId, sheetName, cellId, data),
                result -> {
                    LOGGER.info("Successfully added data to cell: {}", safeString(cellId));
                    return Map.of("Cell Information", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getCellDataExtensionResponse(String driveId, String fileId, String sheetName, String cellId, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.cellData");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createCellDataValidationMap(driveId, fileId, sheetName, cellId);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createCellDataOperationTags(sheetName, cellId);

        return handler.execute(
                "cellData",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, CELL_ID, cellId),
                SubCatalogConstants.CONFIRM_REENTER_CELL_DATA,
                stateId -> StateMapperUtil.cellDataMetaToMap(driveId, fileId, sheetName, cellId, stateId),
                () -> worksheetService.getCellData(driveId, fileId, sheetName, cellId),
                result -> {
                    LOGGER.info("Successfully retrieved data for cell: {}", safeString(cellId));
                    return Map.of("Cell Information", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getAddMultipleCellDataExtensionResponse(String driveId, String fileId, String sheetName,
                                                                     List<Map<String, Object>> cellIDData, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addMultipleCellData");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddMultipleCellDataValidationMap(driveId, fileId, sheetName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddMultipleCellDataOperationTags(sheetName, cellIDData);

        return handler.execute(
                "addMultipleCellData",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, SHEET_NAME, sheetName, CELL_ID_DATA, cellIDData),
                SubCatalogConstants.CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA,
                stateId -> StateMapperUtil.addMultipleCellDataMetaToMap(driveId, fileId, sheetName, cellIDData, stateId),
                () -> {
                    List<FreeForm> allCellData = new ArrayList<>();

                    for (Map eachMap : cellIDData) {
                        String cellId = (String) eachMap.get(CELL_ID);
                        if (!Worksheets.isValidCellId(cellId)) {
                            throw new IllegalArgumentException(Constants.INVALID_CELL_ID_IS_GIVEN + cellId);
                        }
                        String dataToBeAdded = (String) eachMap.get(Constants.DATA);
                        allCellData.add(worksheetService.addDataToCell(driveId, fileId, sheetName, cellId, dataToBeAdded));
                    }

                    return allCellData;
                },
                result -> {
                    LOGGER.info("Successfully added data to multiple cells");
                    return Map.of("Cell Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

}
