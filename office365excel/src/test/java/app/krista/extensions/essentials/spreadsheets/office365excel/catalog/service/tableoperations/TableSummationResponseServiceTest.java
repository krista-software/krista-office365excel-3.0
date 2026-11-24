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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.ksdk.context.RequestContext;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("TableSummationResponseService Tests")
class TableSummationResponseServiceTest {

    @Mock
    private WorkbookService workbookService;

    @Mock
    private ValidationOrchestrator validationOrchestrator;

    @Mock
    private ErrorHandlingStateManager internalStateManager;

    @Mock
    private ExtensionResponseGenerator responseGenerator;

    @Mock
    private RequestContext requestContext;

    @Mock
    private TelemetryHelper telemetryHelper;

    private TableSummationResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String TABLE_NAME = "Table1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new TableSummationResponseService(workbookService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getTableSummationExtensionResponse returns non-null response and increments telemetry")
    void testGetTableSummationExtensionResponse() {
        List<String> columns = List.of("col1", "col2");

        ExtensionResponse result = service.getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableSummation");
    }

    @Test
    @DisplayName("getTableSummationExtensionResponse with null columns")
    void testGetTableSummationExtensionResponseNullColumns() {
        ExtensionResponse result = service.getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableSummation");
    }

    @Test
    @DisplayName("getTableSummationExtensionResponse with allowRetry false")
    void testGetTableSummationExtensionResponseNoRetry() {
        List<String> columns = List.of("Amount");

        ExtensionResponse result = service.getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, columns, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableSummation");
    }

    @Test
    @DisplayName("getTableSummationExtensionResponse with single column")
    void testGetTableSummationExtensionResponseSingleColumn() {
        List<String> columns = List.of("Revenue");

        ExtensionResponse result = service.getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableSummation");
    }

    @Test
    @DisplayName("getTableSummationExtensionResponse with multiple columns")
    void testGetTableSummationExtensionResponseMultipleColumns() {
        List<String> columns = List.of("col1", "col2", "col3", "col4");

        ExtensionResponse result = service.getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableSummation");
    }
}
