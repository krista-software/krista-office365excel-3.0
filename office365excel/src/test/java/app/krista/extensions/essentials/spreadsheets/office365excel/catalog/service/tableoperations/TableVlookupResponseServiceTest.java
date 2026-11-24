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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.ksdk.context.RequestContext;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("TableVlookupResponseService Tests")
class TableVlookupResponseServiceTest {

    @Mock
    private TableService tableService;

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

    private TableVlookupResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String TABLE_NAME = "Table1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new TableVlookupResponseService(tableService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getTableVlookupExtensionResponse returns non-null response and increments telemetry")
    void testGetTableVlookupExtensionResponse() {
        String lookupValue = "test";
        Double columnIndex = 2.0;
        Boolean exactMatch = true;

        ExtensionResponse result = service.getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableVlookup");
    }

    @Test
    @DisplayName("getTableVlookupExtensionResponse with exactMatch false")
    void testGetTableVlookupExtensionResponseExactMatchFalse() {
        String lookupValue = "approxValue";
        Double columnIndex = 3.0;
        Boolean exactMatch = false;

        ExtensionResponse result = service.getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableVlookup");
    }

    @Test
    @DisplayName("getTableVlookupExtensionResponse with null exactMatch")
    void testGetTableVlookupExtensionResponseNullExactMatch() {
        String lookupValue = "value";
        Double columnIndex = 1.0;

        ExtensionResponse result = service.getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableVlookup");
    }

    @Test
    @DisplayName("getTableVlookupExtensionResponse with allowRetry false")
    void testGetTableVlookupExtensionResponseNoRetry() {
        String lookupValue = "searchMe";
        Double columnIndex = 4.0;
        Boolean exactMatch = true;

        ExtensionResponse result = service.getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableVlookup");
    }

    @Test
    @DisplayName("getTableVlookupExtensionResponse with numeric lookup value")
    void testGetTableVlookupExtensionResponseNumericValue() {
        String lookupValue = "42";
        Double columnIndex = 2.0;
        Boolean exactMatch = true;

        ExtensionResponse result = service.getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.tableVlookup");
    }
}
