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
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
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
@DisplayName("SharedSheetResponseService Tests")
class SharedSheetResponseServiceTest {

    @Mock
    private WorkbookService workbookService;

    @Mock
    private WorksheetService worksheetService;

    @Mock
    private WorksheetOperation worksheetOperation;

    @Mock
    private RestClientProvider restClientProvider;

    @Mock
    private SheetFormattingService sheetFormattingService;

    @Mock
    private StringUtil stringUtil;

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

    private SharedSheetResponseService service;

    private static final String DRIVE_ID = "driveId";
    private static final String FILE_ID = "fileId";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        service = new SharedSheetResponseService(workbookService, worksheetService, worksheetOperation,
                restClientProvider, sheetFormattingService, stringUtil, validationOrchestrator,
                internalStateManager, responseGenerator, requestContext, telemetryHelper);
    }

    @Test
    @DisplayName("getSheetExtensionResponse returns non-null response and increments telemetry")
    void testGetSheetExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromFile = null;

        ExtensionResponse result = service.getSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromFile, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsToSheet");
    }

    @Test
    @DisplayName("getRangeSumExtensionResponse returns non-null response and increments telemetry")
    void testGetRangeSumExtensionResponse() {
        String range = "A1:A10";

        ExtensionResponse result = service.getRangeSumExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.rangeSum");
    }

    @Test
    @DisplayName("getSheetVlookupExtensionResponse returns non-null response and increments telemetry")
    void testGetSheetVlookupExtensionResponse() {
        String lookupValue = "test";
        String range = "A1:B10";
        Double columnIndex = 2.0;
        Boolean exactMatch = true;

        ExtensionResponse result = service.getSheetVlookupExtensionResponse(DRIVE_ID, FILE_ID, lookupValue,
                SHEET_NAME, range, columnIndex, exactMatch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.sheetVlookup");
    }

    @Test
    @DisplayName("getRowsForRangeExtensionResponse returns non-null response and increments telemetry")
    void testGetRowsForRangeExtensionResponse() {
        String range = "A1:C5";

        ExtensionResponse result = service.getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, range, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getRowsForRange");
    }

    @Test
    @DisplayName("getSearchExtensionResponse returns non-null response and increments telemetry")
    void testGetSearchExtensionResponse() {
        String dataToSearch = "findMe";

        ExtensionResponse result = service.getSearchExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, dataToSearch, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.search");
    }

    @Test
    @DisplayName("getDeleteColumnWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetDeleteColumnWithIndexExtensionResponse() {
        String columnName = "B";

        ExtensionResponse result = service.getDeleteColumnWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, columnName, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteColumnWithIndex");
    }

    @Test
    @DisplayName("getAddRowWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetAddRowWithIndexExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsv = null;
        Double rowIndex = 3.0;

        ExtensionResponse result = service.getAddRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rows, rowsFromCsv, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowWithIndex");
    }

    @Test
    @DisplayName("getUpdateRowWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetUpdateRowWithIndexExtensionResponse() {
        String row = "updatedRow";
        Double rowIndex = 2.0;

        ExtensionResponse result = service.getUpdateRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, row, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.updateRowWithIndex");
    }

    @Test
    @DisplayName("getCompleteRowInfoExtensionResponse returns non-null response and increments telemetry")
    void testGetCompleteRowInfoExtensionResponse() {
        Double rowIndex = 5.0;

        ExtensionResponse result = service.getCompleteRowInfoExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getCompleteRowInfo");
    }

    @Test
    @DisplayName("getDeleteRowWithIndexExtensionResponse returns non-null response and increments telemetry")
    void testGetDeleteRowWithIndexExtensionResponse() {
        Double rowIndex = 3.0;

        ExtensionResponse result = service.getDeleteRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.deleteRowWithIndex");
    }

    @Test
    @DisplayName("getColumnValuesForKeysInSheetExtensionResponse returns non-null response and increments telemetry")
    void testGetColumnValuesForKeysInSheetExtensionResponse() {
        String range = "A1:D10";
        List<String> lookupRows = List.of("key1");
        List<String> lookupColumns = List.of("col1");
        Boolean getAllRecords = false;
        Double columnNumber = 1.0;

        ExtensionResponse result = service.getColumnValuesForKeysInSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME,
                range, lookupRows, lookupColumns, getAllRecords, columnNumber, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.getColumnValuesForKeysInSheet");
    }

    @Test
    @DisplayName("getClearSharedSheetForRangeExtensionResponse returns non-null response and increments telemetry")
    void testGetClearSharedSheetForRangeExtensionResponse() {
        String cellRange = "A1:Z100";

        ExtensionResponse result = service.getClearSharedSheetForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, cellRange, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.clearSharedSheetForRange");
    }

    @Test
    @DisplayName("getSheetExtensionResponse with null rows and rowsFromFile")
    void testGetSheetExtensionResponseNullRows() {
        ExtensionResponse result = service.getSheetExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, null, null, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowsToSheet");
    }

    @Test
    @DisplayName("getAddRowWithIndexExtensionResponse with null rows")
    void testGetAddRowWithIndexExtensionResponseNullRows() {
        Double rowIndex = 1.0;

        ExtensionResponse result = service.getAddRowWithIndexExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, null, null, rowIndex, ALLOW_RETRY);

        assertNotNull(result);
        verify(telemetryHelper).incrementCount("office365Excel.addRowWithIndex");
    }
}
