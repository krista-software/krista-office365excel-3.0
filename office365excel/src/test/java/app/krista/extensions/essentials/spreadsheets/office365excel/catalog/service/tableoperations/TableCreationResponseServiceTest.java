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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
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
@DisplayName("TableCreationResponseService Tests")
class TableCreationResponseServiceTest {

    @Mock
    private TableService tableService;

    @Mock
    private WorksheetOperation worksheetOperation;

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

    private TableCreationResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final String TABLE_NAME = "Table1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new TableCreationResponseService(tableService, worksheetOperation, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getCreateTableForRangeExtensionResponse returns non-null response and increments telemetry")
    void testGetCreateTableForRangeExtensionResponse() {
        String range = "A1:D10";
        String driveId = "driveId";

        ExtensionResponse result = service.getCreateTableForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, driveId, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.createTableForRange");
    }

    @Test
    @DisplayName("getCreateTableForRangeExtensionResponse with null driveId")
    void testGetCreateTableForRangeExtensionResponseNullDriveId() {
        String range = "A1:D10";

        ExtensionResponse result = service.getCreateTableForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.createTableForRange");
    }

    @Test
    @DisplayName("getColumnValuesForKeysExtensionResponse returns non-null response and increments telemetry")
    void testGetColumnValuesForKeysExtensionResponse() {
        List<String> keyRows = List.of("key1");
        List<String> lookupColumns = List.of("col1");
        Boolean getAllRecord = false;
        Double columnNumber = 1.0;

        ExtensionResponse result = service.getColumnValuesForKeysExtensionResponse(FILE_ID, TABLE_NAME, keyRows, lookupColumns, getAllRecord, columnNumber, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeys");
    }

    @Test
    @DisplayName("getColumnValuesForKeysExtensionResponse with null keyRows and lookupColumns")
    void testGetColumnValuesForKeysExtensionResponseNullLists() {
        Boolean getAllRecord = true;
        Double columnNumber = null;

        ExtensionResponse result = service.getColumnValuesForKeysExtensionResponse(FILE_ID, TABLE_NAME, null, null, getAllRecord, columnNumber, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeys");
    }

    @Test
    @DisplayName("getCreateTableForRangeExtensionResponse with allowRetry false")
    void testGetCreateTableForRangeExtensionResponseNoRetry() {
        String range = "B2:E20";
        String driveId = "driveId";

        ExtensionResponse result = service.getCreateTableForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, driveId, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.createTableForRange");
    }

    @Test
    @DisplayName("getColumnValuesForKeysExtensionResponse with getAllRecord true")
    void testGetColumnValuesForKeysExtensionResponseGetAll() {
        List<String> keyRows = List.of("key1", "key2");
        List<String> lookupColumns = List.of("col1", "col2");
        Boolean getAllRecord = true;
        Double columnNumber = 2.0;

        ExtensionResponse result = service.getColumnValuesForKeysExtensionResponse(FILE_ID, TABLE_NAME, keyRows, lookupColumns, getAllRecord, columnNumber, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeys");
    }
}
