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

@Service
public class CellResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(CellResponseService.class);

    private final WorksheetService worksheetService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    private final WorkbookService workbookService;

    @Inject
    public CellResponseService(WorksheetService worksheetService, ValidationOrchestrator validationOrchestrator,
                               ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                               RequestContext requestContext, TelemetryHelper telemetryHelper, WorkbookService workbookService) {
        this.worksheetService = worksheetService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
        this.workbookService = workbookService;
    }

    public ExtensionResponse getGetCellValueExtensionResponse(String fileId, String sheetName, String cellId, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getCellValue");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetCellValueValidationMap(fileId, sheetName, cellId);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetCellValueOperationTags(fileId, sheetName, cellId);

        return handler.execute(
                "getCellValue",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, CELL_ID, cellId),
                SubCatalogConstants.CONFIRM_REENTER_GET_CELL_VALUE,
                stateId -> StateMapperUtil.getCellValueMetaToMap(fileId, sheetName, cellId, stateId),
                () -> this.worksheetService.getCellData(null, fileId, sheetName, cellId),
                result -> {
                    LOGGER.info("Successfully retrieved cell value for cell: {}", safeString(cellId));
                    return Map.of("Cell Information", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getAddDataToCellExtensionResponse(String fileID, String sheetName, String cellID, String data, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addDataToCell");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddDataToCellValidationMap(fileID, sheetName, cellID);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddDataToCellOperationTags(cellID, data);

        return handler.execute(
                "addDataToCell",
                validationMap,
                Map.of(FILE_ID, fileID, SHEET_NAME, sheetName, CELL_ID, cellID, DATA, data),
                SubCatalogConstants.CONFIRM_RETRY_ADD_DATA_TO_CELL,
                stateId -> StateMapperUtil.addDataToCellMetaToMap(null, fileID, sheetName, cellID, data, stateId),
                () -> this.worksheetService.addDataToCell(null, fileID, sheetName, cellID, data),
                result -> {
                    LOGGER.info("Successfully added data to cell: {}", safeString(cellID));
                    return Map.of("Cell Information", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getAddMultipleCellDataExtensionResponse(String fileId, String sheetName, List<Map<String, Object>> cellIDData, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addMultipleCellData");

        AreaOperationHandler<List<FreeForm>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddMultipleCellDataValidationMapForSheet(fileId, sheetName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddMultipleCellDataOperationTags(sheetName, cellIDData);

        return handler.execute(
                "addMultipleCellData",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, CELL_ID_DATA, cellIDData),
                SubCatalogConstants.CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET,
                stateId -> StateMapperUtil.addMultipleCellDataMetaToMapForSheet(fileId, sheetName, cellIDData, stateId),
                () -> processMultipleCellData(fileId, sheetName, cellIDData),
                result -> {
                    LOGGER.info("Successfully added data to multiple cells");
                    return Map.of("Cell Data", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getAddFormulaToCellExtensionResponse(String fileID, String sheetName, String cellID, String formula, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.addFormulaToCell");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createAddFormulaToCellValidationMap(fileID, sheetName, cellID, formula);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createAddFormulaToCellOperationTags(cellID, formula);

        return handler.execute(
                "addFormulaToCell",
                validationMap,
                Map.of(FILE_ID, fileID, SHEET_NAME, sheetName, CELL_ID, cellID, FORMULA, formula),
                SubCatalogConstants.CONFIRM_REENTER_ADD_FORMULA_TO_CELL,
                stateId -> StateMapperUtil.addFormulaToCellMetaToMap(fileID, sheetName, cellID, formula, stateId),
                () -> workbookService.addFormula(fileID, sheetName, cellID, formula),
                result -> {
                    LOGGER.info("Successfully added formula to cell: {}", safeString(cellID));
                    return Map.of("Cell Information", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    private List<FreeForm> processMultipleCellData(String fileId, String sheetName, List<Map<String, Object>> cellIDData) {
        List<FreeForm> allCellData = new ArrayList<>();

        for (Map eachMap : cellIDData) {
            String cellId = (String) eachMap.get(CELL_ID);
            if (!Worksheets.isValidCellId(cellId)) {
                throw new IllegalArgumentException(Constants.INVALID_CELL_ID_IS_GIVEN + cellId);
            }
            String dataToBeAdded = (String) eachMap.get(Constants.DATA);
            allCellData.add(worksheetService.addDataToCell(null, fileId, sheetName, cellId, dataToBeAdded));
        }

        return allCellData;
    }

}
