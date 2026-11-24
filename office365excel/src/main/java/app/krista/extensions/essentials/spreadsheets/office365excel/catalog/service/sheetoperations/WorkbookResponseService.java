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
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class WorkbookResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookResponseService.class);

    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    private final WorkbookService workbookService;

    @Inject
    public WorkbookResponseService(ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager,
                                   ExtensionResponseGenerator responseGenerator, RequestContext requestContext,
                                   TelemetryHelper telemetryHelper, WorkbookService workbookService) {
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
        this.workbookService = workbookService;
    }

    public ExtensionResponse getGetExcelFileInformationExtensionResponse(String path, String name, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getExcelFileInformation");

        AreaOperationHandler<ExcelFile> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createGetExcelFileInfoValidationMap(path, name);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createGetExcelFileInfoOperationTags(path, name);

        return handler.execute(
                "getExcelFileInformation",
                validationMap,
                Map.of(PATH, path, NAME, name),
                SubCatalogConstants.CONFIRM_REENTER_GET_EXCEL_FILE_INFO,
                stateId -> StateMapperUtil.getExcelFileInfoMetaToMap(path, name, stateId),
                () -> workbookService.getWorkbookDetails(path, name),
                result -> {
                    LOGGER.info("Successfully retrieved Excel file information for: {}", safeString(name));
                    return Map.of("Excel File", result);
                },
                telemetryTags,
                allowRetry
        );
    }

}
