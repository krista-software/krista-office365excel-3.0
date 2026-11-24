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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

/**
 * Helper class for creating extension responses for shared file operations
 */
@Service
public class SharedFileResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(SharedFileResponseService.class);

    private final WorkbookService workbookService;
    private final DriveService driveService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;


    @Inject
    public SharedFileResponseService(WorkbookService workbookService, DriveService driveService, ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator, RequestContext requestContext, TelemetryHelper telemetryHelper) {
        this.workbookService = workbookService;
        this.driveService = driveService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    public ExtensionResponse getSharedExcelFileExtensionResponse(String sharedFolderName, String sharedFileName, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getSharedExcelFile");

        AreaOperationHandler<ExcelFile> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetSharedExcelFileValidationMap(sharedFileName);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetSharedExcelFileOperationTags(sharedFolderName, sharedFileName);

        return handler.execute(
                "getSharedExcelFile",
                validationMap,
                Map.of(SHARED_FOLDER_NAME, sharedFolderName != null ? sharedFolderName : "", SHARED_FILE_NAME, sharedFileName),
                SubCatalogConstants.CONFIRM_REENTER_GET_SHARED_EXCEL_FILE,
                stateId -> StateMapperUtil.getSharedExcelFileMetaToMap(sharedFolderName, sharedFileName, stateId),
                () -> {
                    List<ExcelFile> excelFileFromSharedDrive = driveService.getExcelFileFromSharedDrive(sharedFolderName, sharedFileName);

                    if (excelFileFromSharedDrive != null && !excelFileFromSharedDrive.isEmpty()) {
                        return excelFileFromSharedDrive.getFirst();
                    } else {
                        throw new RuntimeException("Excel file not found: " + sharedFileName);
                    }
                },
                result -> {
                    LOGGER.info("Successfully retrieved shared Excel file: {}", safeString(sharedFileName));
                    return Map.of("Shared Excel File", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getConvertToPdfExtensionResponse(String driveId, String fileId, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.convertToPDF");

        AreaOperationHandler<File> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createConvertToPdfValidationMap(driveId, fileId);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createConvertToPdfOperationTags(driveId, fileId);

        return handler.execute(
                "convertToPDF",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId),
                SubCatalogConstants.CONFIRM_REENTER_CONVERT_TO_PDF,
                stateId -> StateMapperUtil.convertToPdfMetaToMap(driveId, fileId, stateId),
                () -> {
                    try {
                        return driveService.getPdfFile(fileId, driveId);
                    } catch (IOException cause) {
                        throw new RuntimeException(cause);
                    }
                },
                result -> {
                    LOGGER.info("Successfully converted Excel file to PDF");
                    return Map.of("PDF File", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getExecuteFormulaExtensionResponse(String driveId, String fileId, String formula, String formulaName, String comment, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.executeFormulaForSharedFile");

        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createExecuteFormulaValidationMap(driveId, fileId, formula, formulaName, comment);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createExecuteFormulaOperationTags(formula, formulaName);

        return handler.execute(
                "executeFormulaForSharedFile",
                validationMap,
                Map.of(DRIVE_ID, driveId, FILE_ID, fileId, FORMULA, formula, FORMULA_NAME, formulaName, COMMENT, comment),
                SubCatalogConstants.CONFIRM_REENTER_EXECUTE_FORMULA,
                stateId -> StateMapperUtil.executeFormulaMetaToMap(driveId, fileId, formula, formulaName, comment, stateId),
                () -> workbookService.executeFormula(driveId, fileId, formula, formulaName, comment),
                result -> {
                    LOGGER.info("Successfully executed formula: {}", safeString(formula));
                    return Map.of("Formula Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }

}
