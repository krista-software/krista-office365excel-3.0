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
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.util.List;

@Service
public class TableOperationsFacade {

    private final TableResponseService tableResponseService;
    private final TableVlookupResponseService tableVlookupResponseService;
    private final TableSummationResponseService tableSummationResponseService;
    private final TableCreationResponseService tableCreationResponseService;

    @Inject
    public TableOperationsFacade(TableResponseService tableResponseService,
                                 TableVlookupResponseService tableVlookupResponseService,
                                 TableSummationResponseService tableSummationResponseService,
                                 TableCreationResponseService tableCreationResponseService) {
        this.tableResponseService = tableResponseService;
        this.tableVlookupResponseService = tableVlookupResponseService;
        this.tableSummationResponseService = tableSummationResponseService;
        this.tableCreationResponseService = tableCreationResponseService;
    }

    // Table CRUD Operations
    public ExtensionResponse getAddRowsToTableExtensionResponse(String fileID, String driveID, String tableName,
                                                                List<String> rows, List<File> rowsFromCsv,
                                                                Double index, String separator, Boolean allowRetry) {
        return tableResponseService.getAddRowsToTableExtensionResponse(fileID, driveID, tableName, rows, rowsFromCsv, index, separator, allowRetry);
    }

    public ExtensionResponse getTableRowsExtensionResponse(String workbookID, String tableName, Boolean allowRetry) {
        return tableResponseService.getTableRowsExtensionResponse(workbookID, tableName, allowRetry);
    }

    public ExtensionResponse getAddRowsExtensionResponse(String fileID, String tableName, List<String> rows, List<File> rowsFromCsv, Boolean allowRetry) {
        return tableResponseService.getAddRowsExtensionResponse(fileID, tableName, rows, rowsFromCsv, allowRetry);
    }

    public ExtensionResponse getDeleteTableColumnExtensionResponse(String fileID, String tableName, String columnName, Boolean allowRetry) {
        return tableResponseService.getDeleteTableColumnExtensionResponse(fileID, tableName, columnName, allowRetry);
    }

    // Table Vlookup Operations
    public ExtensionResponse getTableVlookupExtensionResponse(String fileId, String lookupValue, String tableName,
                                                              Double columnIndex, Boolean exactMatch, Boolean allowRetry) {
        return tableVlookupResponseService.getTableVlookupExtensionResponse(fileId, lookupValue, tableName, columnIndex, exactMatch, allowRetry);
    }

    // Table Summation Operations
    public ExtensionResponse getTableSummationExtensionResponse(String fileID, String tableName, List<String> columns, Boolean allowRetry) {
        return tableSummationResponseService.getTableSummationExtensionResponse(fileID, tableName, columns, allowRetry);
    }

    // Table Creation Operations
    public ExtensionResponse getCreateTableForRangeExtensionResponse(String fileId, String sheetName, String range, String driveId, Boolean allowRetry) {
        return tableCreationResponseService.getCreateTableForRangeExtensionResponse(fileId, sheetName, range, driveId, allowRetry);
    }

    public ExtensionResponse getColumnValuesForKeysExtensionResponse(String fileID, String tableName, List<String> keyRows,
                                                                     List<String> lookupColumns, Boolean getAllRecord, Double columnNumber, Boolean allowRetry) {
        return tableCreationResponseService.getColumnValuesForKeysExtensionResponse(fileID, tableName, keyRows, lookupColumns, getAllRecord, columnNumber, allowRetry);
    }
}