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
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class SearchResponseService {

    private static final Logger LOGGER = LoggerFactory.getLogger(SearchResponseService.class);

    private final WorksheetService worksheetService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public SearchResponseService(WorksheetService worksheetService, ValidationOrchestrator validationOrchestrator,
                                 ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                 RequestContext requestContext, TelemetryHelper telemetryHelper) {
        this.worksheetService = worksheetService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    public ExtensionResponse getFindValueExtensionResponse(String fileId, String sheetName, String lookupValue, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.findValue");

        AreaOperationHandler<List<String>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createFindValueValidationMapForSheet(fileId, sheetName, lookupValue);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createFindValueOperationTags(sheetName, lookupValue);

        return handler.execute(
                "findValue",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, LOOKUP_VALUE, lookupValue),
                SubCatalogConstants.CONFIRM_REENTER_FIND_VALUE,
                stateId -> StateMapperUtil.findValueMetaToMap(fileId, sheetName, lookupValue, stateId),
                () -> worksheetService.findValue(null, fileId, sheetName, lookupValue),
                result -> {
                    LOGGER.info("Successfully found value: {} in sheet: {}", safeString(lookupValue), safeString(sheetName));
                    return Map.of("Cell IDs", result);
                },
                telemetryTags,
                allowRetry
        );
    }

    public ExtensionResponse getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(String fileId, String sheetName, String keyColumnName, String columnRange, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.getMinMaxValueFromRowDataByColumnRange");

        AreaOperationHandler<Map<String, Object>> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createMinMaxValueValidationMap(fileId, sheetName, keyColumnName, columnRange);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createMinMaxValueOperationTags(keyColumnName, columnRange);

        return handler.execute(
                "getMinMaxValueFromRowDataByColumnRange",
                validationMap,
                Map.of(FILE_ID, fileId, SHEET_NAME, sheetName, KEY_COLUMN_NAME, keyColumnName, COLUMN_RANGE, columnRange),
                SubCatalogConstants.CONFIRM_REENTER_GET_MIN_MAX_VALUE,
                stateId -> StateMapperUtil.getMinMaxValueMetaToMap(fileId, sheetName, keyColumnName, columnRange, stateId),
                () -> this.worksheetService.findMinMaxValueFromRowDataByColumnRange(null, fileId, sheetName, keyColumnName, columnRange),
                result -> {
                    LOGGER.info("Successfully retrieved min-max values for column range: {}", safeString(columnRange));
                    return Map.of("Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }
}
