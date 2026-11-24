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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.TableService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
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
@DisplayName("SharedTableResponseService Tests")
class SharedTableResponseServiceTest {

    @Mock
    private WorkbookService workbookService;

    @Mock
    private WorksheetService worksheetService;

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

    private SharedTableResponseService service;

    private static final String DRIVE_ID = "driveId";
    private static final String FILE_ID = "fileId";
    private static final String TABLE_NAME = "Table1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new SharedTableResponseService(workbookService, worksheetService, tableService,
                worksheetOperation, validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getExtensionResponse returns non-null response and increments telemetry")
    void testGetExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsvFile = null;

        ExtensionResponse result = service.getExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, rows, rowsFromCsvFile, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsToTable");
    }

    @Test
    @DisplayName("getExtensionResponse with null rows")
    void testGetExtensionResponseNullRows() {
        ExtensionResponse result = service.getExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, null, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsToTable");
    }

    @Test
    @DisplayName("getTableRowsExtensionResponse returns non-null response and increments telemetry")
    void testGetTableRowsExtensionResponse() {
        ExtensionResponse result = service.getTableRowsExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getTableRows");
    }

    @Test
    @DisplayName("getDeleteColumnExtensionResponse returns non-null response and increments telemetry")
    void testGetDeleteColumnExtensionResponse() {
        String columnName = "ColumnA";

        ExtensionResponse result = service.getDeleteColumnExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteColumn");
    }

    @Test
    @DisplayName("getColumnValuesForKeysExtensionResponse returns non-null response and increments telemetry")
    void testGetColumnValuesForKeysExtensionResponse() {
        List<String> keyRows = List.of("key1");
        List<String> lookupColumn = List.of("col1");
        Boolean getAllRecord = true;

        ExtensionResponse result = service.getColumnValuesForKeysExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, keyRows, lookupColumn, getAllRecord, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeys");
    }

    @Test
    @DisplayName("getVlookupInTableExtensionResponse returns non-null response and increments telemetry")
    void testGetVlookupInTableExtensionResponse() {
        String lookupValue = "testValue";
        Double columnIndex = 2.0;
        Boolean exactMatch = false;

        ExtensionResponse result = service.getVlookupInTableExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.vlookupInTable");
    }

    @Test
    @DisplayName("getColumnsSumExtensionResponse returns non-null response and increments telemetry")
    void testGetColumnsSumExtensionResponse() {
        List<String> columns = List.of("col1", "col2");

        ExtensionResponse result = service.getColumnsSumExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.columnsSum");
    }

    @Test
    @DisplayName("getTableRowsExtensionResponse with allowRetry false")
    void testGetTableRowsExtensionResponseNoRetry() {
        ExtensionResponse result = service.getTableRowsExtensionResponse(DRIVE_ID, FILE_ID, TABLE_NAME, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getTableRows");
    }

    @Test
    @DisplayName("getVlookupInTableExtensionResponse with exact match true")
    void testGetVlookupInTableExtensionResponseExactMatch() {
        String lookupValue = "exactValue";
        Double columnIndex = 1.0;
        Boolean exactMatch = true;

        ExtensionResponse result = service.getVlookupInTableExtensionResponse(DRIVE_ID, FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.vlookupInTable");
    }
}
