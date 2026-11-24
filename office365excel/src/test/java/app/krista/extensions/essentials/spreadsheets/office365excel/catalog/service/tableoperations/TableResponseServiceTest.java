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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
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
@DisplayName("TableResponseService Tests")
class TableResponseServiceTest {

    @Mock
    private WorkbookService workbookService;

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

    @Mock
    private WorksheetService worksheetService;

    private TableResponseService service;

    private static final String FILE_ID = "fileId";
    private static final String DRIVE_ID = "driveId";
    private static final String TABLE_NAME = "Table1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new TableResponseService(workbookService, tableService, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper, worksheetService);
    }

    @Test
    @DisplayName("getAddRowsToTableExtensionResponse returns non-null response and increments telemetry")
    void testGetAddRowsToTableExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsv = null;
        Double index = 1.0;
        String separator = ",";

        ExtensionResponse result = service.getAddRowsToTableExtensionResponse(FILE_ID, DRIVE_ID, TABLE_NAME, rows, rowsFromCsv, index, separator, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsToTableWithIndex");
    }

    @Test
    @DisplayName("getAddRowsToTableExtensionResponse with null rows")
    void testGetAddRowsToTableExtensionResponseNullRows() {
        Double index = 2.0;
        String separator = ",";

        ExtensionResponse result = service.getAddRowsToTableExtensionResponse(FILE_ID, DRIVE_ID, TABLE_NAME, null, null, index, separator, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsToTableWithIndex");
    }

    @Test
    @DisplayName("getTableRowsExtensionResponse returns non-null response and increments telemetry")
    void testGetTableRowsExtensionResponse() {
        String workbookID = "workbookId";

        ExtensionResponse result = service.getTableRowsExtensionResponse(workbookID, TABLE_NAME, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getTableRows");
    }

    @Test
    @DisplayName("getAddRowsExtensionResponse returns non-null response and increments telemetry")
    void testGetAddRowsExtensionResponse() {
        List<String> rows = List.of("row1", "row2");
        List<File> rowsFromCsv = null;

        ExtensionResponse result = service.getAddRowsExtensionResponse(FILE_ID, TABLE_NAME, rows, rowsFromCsv, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsInTable");
    }

    @Test
    @DisplayName("getAddRowsExtensionResponse with null rows")
    void testGetAddRowsExtensionResponseNullRows() {
        ExtensionResponse result = service.getAddRowsExtensionResponse(FILE_ID, TABLE_NAME, null, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsInTable");
    }

    @Test
    @DisplayName("getDeleteTableColumnExtensionResponse returns non-null response and increments telemetry")
    void testGetDeleteTableColumnExtensionResponse() {
        String columnName = "ColumnA";

        ExtensionResponse result = service.getDeleteTableColumnExtensionResponse(FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteTableColumn");
    }

    @Test
    @DisplayName("getTableRowsExtensionResponse with allowRetry false")
    void testGetTableRowsExtensionResponseNoRetry() {
        ExtensionResponse result = service.getTableRowsExtensionResponse(FILE_ID, TABLE_NAME, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getTableRows");
    }

    @Test
    @DisplayName("getDeleteTableColumnExtensionResponse with allowRetry false")
    void testGetDeleteTableColumnExtensionResponseNoRetry() {
        String columnName = "ColB";

        ExtensionResponse result = service.getDeleteTableColumnExtensionResponse(FILE_ID, TABLE_NAME, columnName, false);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteTableColumn");
    }

    @Test
    @DisplayName("getAddRowsToTableExtensionResponse with different separator")
    void testGetAddRowsToTableExtensionResponseDifferentSeparator() {
        List<String> rows = List.of("row1");
        Double index = 1.0;
        String separator = ";";

        ExtensionResponse result = service.getAddRowsToTableExtensionResponse(FILE_ID, DRIVE_ID, TABLE_NAME, rows, null, index, separator, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsToTableWithIndex");
    }
}
