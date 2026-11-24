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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.tableoperations.TableCreationResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.tableoperations.TableResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.tableoperations.TableSummationResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.tableoperations.TableVlookupResponseService;
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
@DisplayName("TableOperationsFacade Tests")
class TableOperationsFacadeTest {

    @Mock
    private TableResponseService tableResponseService;

    @Mock
    private TableVlookupResponseService tableVlookupResponseService;

    @Mock
    private TableSummationResponseService tableSummationResponseService;

    @Mock
    private TableCreationResponseService tableCreationResponseService;

    @Mock
    private ExtensionResponse mockResponse;

    private TableOperationsFacade facade;

    private static final String FILE_ID = "fileId";
    private static final String DRIVE_ID = "driveId";
    private static final String TABLE_NAME = "Table1";
    private static final String SHEET_NAME = "Sheet1";
    private static final Boolean ALLOW_RETRY = true;

    @BeforeEach
    void setUp() {
        facade = new TableOperationsFacade(tableResponseService, tableVlookupResponseService,
                tableSummationResponseService, tableCreationResponseService);
    }

    // Table CRUD Operations

    @Test
    @DisplayName("getAddRowsToTableExtensionResponse delegates to TableResponseService")
    void testGetAddRowsToTableExtensionResponse() {
        List<String> rows = List.of("row1");
        List<File> rowsFromCsv = null;
        Double index = 1.0;
        String separator = ",";
        when(tableResponseService.getAddRowsToTableExtensionResponse(FILE_ID, DRIVE_ID, TABLE_NAME, rows, rowsFromCsv, index, separator, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddRowsToTableExtensionResponse(FILE_ID, DRIVE_ID, TABLE_NAME, rows, rowsFromCsv, index, separator, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableResponseService).getAddRowsToTableExtensionResponse(FILE_ID, DRIVE_ID, TABLE_NAME, rows, rowsFromCsv, index, separator, ALLOW_RETRY);
        verifyNoInteractions(tableVlookupResponseService, tableSummationResponseService, tableCreationResponseService);
    }

    @Test
    @DisplayName("getTableRowsExtensionResponse delegates to TableResponseService")
    void testGetTableRowsExtensionResponse() {
        String workbookID = "workbookId";
        when(tableResponseService.getTableRowsExtensionResponse(workbookID, TABLE_NAME, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getTableRowsExtensionResponse(workbookID, TABLE_NAME, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableResponseService).getTableRowsExtensionResponse(workbookID, TABLE_NAME, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getAddRowsExtensionResponse delegates to TableResponseService")
    void testGetAddRowsExtensionResponse() {
        List<String> rows = List.of("row1", "row2");
        List<File> rowsFromCsv = null;
        when(tableResponseService.getAddRowsExtensionResponse(FILE_ID, TABLE_NAME, rows, rowsFromCsv, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getAddRowsExtensionResponse(FILE_ID, TABLE_NAME, rows, rowsFromCsv, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableResponseService).getAddRowsExtensionResponse(FILE_ID, TABLE_NAME, rows, rowsFromCsv, ALLOW_RETRY);
    }

    @Test
    @DisplayName("getDeleteTableColumnExtensionResponse delegates to TableResponseService")
    void testGetDeleteTableColumnExtensionResponse() {
        String columnName = "ColumnA";
        when(tableResponseService.getDeleteTableColumnExtensionResponse(FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getDeleteTableColumnExtensionResponse(FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableResponseService).getDeleteTableColumnExtensionResponse(FILE_ID, TABLE_NAME, columnName, ALLOW_RETRY);
    }

    // Table Vlookup Operations

    @Test
    @DisplayName("getTableVlookupExtensionResponse delegates to TableVlookupResponseService")
    void testGetTableVlookupExtensionResponse() {
        String lookupValue = "test";
        Double columnIndex = 2.0;
        Boolean exactMatch = true;
        when(tableVlookupResponseService.getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableVlookupResponseService).getTableVlookupExtensionResponse(FILE_ID, lookupValue, TABLE_NAME, columnIndex, exactMatch, ALLOW_RETRY);
        verifyNoInteractions(tableResponseService, tableSummationResponseService, tableCreationResponseService);
    }

    // Table Summation Operations

    @Test
    @DisplayName("getTableSummationExtensionResponse delegates to TableSummationResponseService")
    void testGetTableSummationExtensionResponse() {
        List<String> columns = List.of("col1", "col2");
        when(tableSummationResponseService.getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, columns, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableSummationResponseService).getTableSummationExtensionResponse(FILE_ID, TABLE_NAME, columns, ALLOW_RETRY);
        verifyNoInteractions(tableResponseService, tableVlookupResponseService, tableCreationResponseService);
    }

    // Table Creation Operations

    @Test
    @DisplayName("getCreateTableForRangeExtensionResponse delegates to TableCreationResponseService")
    void testGetCreateTableForRangeExtensionResponse() {
        String range = "A1:D10";
        when(tableCreationResponseService.getCreateTableForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, DRIVE_ID, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getCreateTableForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, DRIVE_ID, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableCreationResponseService).getCreateTableForRangeExtensionResponse(FILE_ID, SHEET_NAME, range, DRIVE_ID, ALLOW_RETRY);
        verifyNoInteractions(tableResponseService, tableVlookupResponseService, tableSummationResponseService);
    }

    @Test
    @DisplayName("getColumnValuesForKeysExtensionResponse delegates to TableCreationResponseService")
    void testGetColumnValuesForKeysExtensionResponse() {
        List<String> keyRows = List.of("key1");
        List<String> lookupColumns = List.of("col1");
        Boolean getAllRecord = false;
        Double columnNumber = 1.0;
        when(tableCreationResponseService.getColumnValuesForKeysExtensionResponse(FILE_ID, TABLE_NAME, keyRows, lookupColumns, getAllRecord, columnNumber, ALLOW_RETRY))
                .thenReturn(mockResponse);

        ExtensionResponse result = facade.getColumnValuesForKeysExtensionResponse(FILE_ID, TABLE_NAME, keyRows, lookupColumns, getAllRecord, columnNumber, ALLOW_RETRY);

        assertSame(mockResponse, result);
        verify(tableCreationResponseService).getColumnValuesForKeysExtensionResponse(FILE_ID, TABLE_NAME, keyRows, lookupColumns, getAllRecord, columnNumber, ALLOW_RETRY);
    }
}
