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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.TableVlookupReentryHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.TelemetryTagsHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.model.base.FreeForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.*;

/**
 * Implementation for table vlookup specific reentry operations
 */
public class TableVlookupReentryHandlerImpl extends BaseReentryHandlerImpl implements TableVlookupReentryHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(TableVlookupReentryHandlerImpl.class);
    private final TableService tableService;
    private final TelemetryHelper telemetryHelper;

    @Inject
    public TableVlookupReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, 
                                         ExtensionResponseGenerator responseGenerator,
                                         TableService tableService,
                                         TelemetryHelper telemetryHelper) {
        super(internalStateManager, responseGenerator);
        this.tableService = tableService;
        this.telemetryHelper = telemetryHelper;
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleTableVlookupConfirmation(Map<String, Object> map) {
        LOGGER.info("TableVlookupReentryHandler.handleTableVlookupConfirmation start: {}", map);
        Boolean reenter = (Boolean) map.get(REENTER);
        String stateId = (String) map.get(STATE_ID);
        Map<String, Object> state = internalStateManager.get(stateId);
        List<ValidationOrchestrator.ValidationResult> validationResults = getValidationResults(state);

        if (Boolean.FALSE.equals(reenter)) {
            LOGGER.info("User chose not to re-enter table vlookup data. Returning fetch deny response.");
            internalStateManager.remove(stateId);
            return responseGenerator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                    validationResults,
                    null,
                    Map.of());
        }

        LOGGER.info(REENTER_WAS_TRUE_HENCE_CONTINUING);
        return responseGenerator.generateFetchResponse(
                ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                validationResults,
                SubCatalogConstants.HANDLE_REENTER_TABLE_VLOOKUP,
                map);
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleTableVlookupReentry(Map<String, Object> map) {
        LOGGER.info("TableVlookupReentryHandler.handleTableVlookupReentry start: {}", map);
        String stateId = (String) map.get(STATE_ID);

        long startTime = System.currentTimeMillis();
        Map<String, String> telemetryTags = TelemetryTagsHelper.createVlookupInTableOperationTags(
                (String) map.get(LOOKUP_VALUE), (String) map.get(TABLE_NAME), (Double) map.get(COLUMN_INDEX));
        telemetryHelper.incrementCount("office365Excel.tableVlookupReentry");

        try {
            String fileId = (String) map.get(FILE_ID);
            String lookupValue = (String) map.get(LOOKUP_VALUE);
            String tableName = (String) map.get(TABLE_NAME);
            Double columnIndex = (Double) map.get(COLUMN_INDEX);
            Boolean exactMatch = (Boolean) map.get(EXACT_MATCH);

            LOGGER.info("Processing corrected input for table vlookup operation");
            FreeForm result = this.tableService.vlookup(null, fileId, lookupValue, tableName, columnIndex, exactMatch);

            LOGGER.info("Successfully performed table vlookup for value: {}", lookupValue);
            return ExtensionResponseFactory.create(Map.of("Lookup Result", result));
        } catch (Exception cause) {
            LOGGER.error("Table vlookup reentry handling failed: {}", cause.getMessage(), cause);
            telemetryHelper.recordError("office365Excel.tableVlookupReentry", startTime, cause, telemetryTags);
            return ExtensionResponseFactory.create(
                    cause,
                    "Failed to perform table vlookup on retry: " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction("An error occurred while performing table vlookup. Please try again.", List.of())),
                    null, null);
        } finally {
            internalStateManager.remove(stateId);
            telemetryHelper.recordSuccess("office365Excel.tableVlookupReentry", startTime, telemetryTags);
        }
    }
}