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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.tableoperations;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.FreeForm;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Service
public class TableVlookupResponseService {
    
    private static final Logger LOGGER = LoggerFactory.getLogger(TableVlookupResponseService.class);
    
    private final TableService tableService;
    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;
    
    @Inject
    public TableVlookupResponseService(TableService tableService, ValidationOrchestrator validationOrchestrator,
                                      ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator,
                                      RequestContext requestContext, TelemetryHelper telemetryHelper) {
        this.tableService = tableService;
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }
    
    public ExtensionResponse getTableVlookupExtensionResponse(String fileId, String lookupValue, String tableName,
                                                             Double columnIndex, Boolean exactMatch, Boolean allowRetry) {
        telemetryHelper.incrementCount("office365Excel.tableVlookup");
        
        AreaOperationHandler<FreeForm> handler = new AreaOperationHandler<>(validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
        Map<Validator.ValidationResource, String> validationMap = ValidationMapHelper.createTableVlookupValidationMap(fileId, lookupValue, tableName, columnIndex);
        Map<String, String> telemetryTags = TelemetryTagsHelper.createTableVlookupOperationTags(lookupValue, tableName, columnIndex);

        return handler.execute(
                "tableVlookup",
                validationMap,
                Map.of(FILE_ID, fileId, LOOKUP_VALUE, lookupValue, TABLE_NAME, tableName, COLUMN_INDEX, columnIndex,
                        EXACT_MATCH, exactMatch != null ? exactMatch : false),
                SubCatalogConstants.CONFIRM_REENTER_TABLE_VLOOKUP,
                stateId -> StateMapperUtil.vlookupTableMetaToMap(fileId, lookupValue, tableName, columnIndex, exactMatch, stateId),
                () -> tableService.vlookup(null, fileId, lookupValue, tableName, columnIndex, exactMatch),
                result -> {
                    LOGGER.info("Successfully performed tableVlookup for value: {} in table: {}", safeString(lookupValue), safeString(tableName));
                    return Map.of("Lookup Result", result);
                },
                telemetryTags,
                allowRetry
        );
    }
}