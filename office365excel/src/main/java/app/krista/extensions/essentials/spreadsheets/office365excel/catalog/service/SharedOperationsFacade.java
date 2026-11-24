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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedCellResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedFileResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedSheetResponseService;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sharedoperations.SharedTableResponseService;
import app.krista.model.base.File;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

@Service
public class SharedOperationsFacade {
    
    private final SharedCellResponseService sharedCellResponseService;
    private final SharedFileResponseService sharedFileResponseService;
    private final SharedSheetResponseService sharedSheetResponseService;
    private final SharedTableResponseService sharedTableResponseService;
    
    @Inject
    public SharedOperationsFacade(SharedCellResponseService sharedCellResponseService,
                                 SharedFileResponseService sharedFileResponseService,
                                 SharedSheetResponseService sharedSheetResponseService,
                                 SharedTableResponseService sharedTableResponseService) {
        this.sharedCellResponseService = sharedCellResponseService;
        this.sharedFileResponseService = sharedFileResponseService;
        this.sharedSheetResponseService = sharedSheetResponseService;
        this.sharedTableResponseService = sharedTableResponseService;
    }
    
    // Shared Cell Operations
    public ExtensionResponse getAddDataToCellExtensionResponse(String driveId, String fileId, String sheetName, String cellId, String data, Boolean allowRetry) {
        return sharedCellResponseService.getAddDataToCellExtensionResponse(driveId, fileId, sheetName, cellId, data, allowRetry);
    }
    
    public ExtensionResponse getCellDataExtensionResponse(String driveId, String fileId, String sheetName, String cellId, Boolean allowRetry) {
        return sharedCellResponseService.getCellDataExtensionResponse(driveId, fileId, sheetName, cellId, allowRetry);
    }
    
    public ExtensionResponse getAddMultipleCellDataExtensionResponse(String driveId, String fileId, String sheetName, List<Map<String, Object>> cellIDData, Boolean allowRetry) {
        return sharedCellResponseService.getAddMultipleCellDataExtensionResponse(driveId, fileId, sheetName, cellIDData, allowRetry);
    }
    
    // Shared File Operations
    public ExtensionResponse getSharedExcelFileExtensionResponse(String sharedFolderName, String sharedFileName, Boolean allowRetry) {
        return sharedFileResponseService.getSharedExcelFileExtensionResponse(sharedFolderName, sharedFileName, allowRetry);
    }
    
    public ExtensionResponse getConvertToPdfExtensionResponse(String driveId, String fileId, Boolean allowRetry) {
        return sharedFileResponseService.getConvertToPdfExtensionResponse(driveId, fileId, allowRetry);
    }
    
    public ExtensionResponse getExecuteFormulaExtensionResponse(String driveId, String fileId, String formula, String formulaName, String comment, Boolean allowRetry) {
        return sharedFileResponseService.getExecuteFormulaExtensionResponse(driveId, fileId, formula, formulaName, comment, allowRetry);
    }
    
    // Shared Sheet Operations
    public ExtensionResponse getSheetExtensionResponse(String driveId, String fileId, String sheetName, List<String> rows, List<File> rowsFromFile, Boolean allowRetry) {
        return sharedSheetResponseService.getSheetExtensionResponse(driveId, fileId, sheetName, rows, rowsFromFile, allowRetry);
    }
    
    public ExtensionResponse getSheetVlookupExtensionResponse(String driveId, String fileId, String lookupValue, String sheetName, String range, Double columnIndex, Boolean exactMatch, Boolean allowRetry) {
        return sharedSheetResponseService.getSheetVlookupExtensionResponse(driveId, fileId, lookupValue, sheetName, range, columnIndex, exactMatch, allowRetry);
    }
    
    public ExtensionResponse getRowsForRangeExtensionResponse(String driveId, String fileId, String sheetName, String range, Boolean allowRetry) {
        return sharedSheetResponseService.getRowsForRangeExtensionResponse(driveId, fileId, sheetName, range, allowRetry);
    }
    
    public ExtensionResponse getSearchExtensionResponse(String driveId, String fileId, String sheetName, String dataToSearch, Boolean allowRetry) {
        return sharedSheetResponseService.getSearchExtensionResponse(driveId, fileId, sheetName, dataToSearch, allowRetry);
    }
    
    public ExtensionResponse getAddRowWithIndexExtensionResponse(String driveId, String fileId, String sheetName, List<String> rows, List<File> rowsFromCsv, Double rowIndex, Boolean allowRetry) {
        return sharedSheetResponseService.getAddRowWithIndexExtensionResponse(driveId, fileId, sheetName, rows, rowsFromCsv, rowIndex, allowRetry);
    }
    
    public ExtensionResponse getUpdateRowWithIndexExtensionResponse(String driveId, String fileId, String sheetName, String row, Double rowIndex, Boolean allowRetry) {
        return sharedSheetResponseService.getUpdateRowWithIndexExtensionResponse(driveId, fileId, sheetName, row, rowIndex, allowRetry);
    }
    
    public ExtensionResponse getColumnValuesForKeysInSheetExtensionResponse(String driveId, String fileId, String sheetName, String range, List<String> lookupRows, List<String> lookupColumns, Boolean getAllRecords, Double columnNumber, Boolean allowRetry) {
        return sharedSheetResponseService.getColumnValuesForKeysInSheetExtensionResponse(driveId, fileId, sheetName, range, lookupRows, lookupColumns, getAllRecords, columnNumber, allowRetry);
    }
    
    public ExtensionResponse getClearSharedSheetForRangeExtensionResponse(String driveId, String fileId, String sheetName, String cellRange, Boolean allowRetry) {
        return sharedSheetResponseService.getClearSharedSheetForRangeExtensionResponse(driveId, fileId, sheetName, cellRange, allowRetry);
    }
    
    // Shared Table Operations
    public ExtensionResponse getExtensionResponse(String driveId, String fileId, String tableName, List<String> rows, List<File> rowsFromCsvFile, Boolean allowRetry) {
        return sharedTableResponseService.getExtensionResponse(driveId, fileId, tableName, rows, rowsFromCsvFile, allowRetry);
    }
    
    public ExtensionResponse getTableRowsExtensionResponse(String driveId, String fileId, String tableName, Boolean allowRetry) {
        return sharedTableResponseService.getTableRowsExtensionResponse(driveId, fileId, tableName, allowRetry);
    }
    
    public ExtensionResponse getColumnValuesForKeysExtensionResponse(String driveId, String fileId, String tableName, List<String> keyRows, List<String> lookupColumn, Boolean getAllRecord, Boolean allowRetry) {
        return sharedTableResponseService.getColumnValuesForKeysExtensionResponse(driveId, fileId, tableName, keyRows, lookupColumn, getAllRecord, allowRetry);
    }

    public ExtensionResponse getRangeSumExtensionResponse(String driveId, String fileId, String sheetName, String range, Boolean allowRetry) {
        return sharedSheetResponseService.getRangeSumExtensionResponse(driveId, fileId, sheetName, range, allowRetry);
    }

    public ExtensionResponse getDeleteColumnExtensionResponse(String driveId, String fileId, String tableName, String columnName, Boolean allowRetry) {
        return sharedTableResponseService.getDeleteColumnExtensionResponse(driveId, fileId, tableName, columnName, allowRetry);
    }

    public ExtensionResponse getVlookupInTableExtensionResponse(String driveId, String fileId, String lookupValue, String tableName, Double columnIndex, Boolean exactMatch, Boolean allowRetry) {
        return sharedTableResponseService.getVlookupInTableExtensionResponse(driveId, fileId, lookupValue, tableName, columnIndex, exactMatch, allowRetry);
    }

    public ExtensionResponse getColumnsSumExtensionResponse(String driveId, String fileId, String tableName, List<String> columns, Boolean allowRetry) {
        return sharedTableResponseService.getColumnsSumExtensionResponse(driveId, fileId, tableName, columns, allowRetry);
    }

    public ExtensionResponse getDeleteColumnWithIndexExtensionResponse(String driveId, String fileId, String sheetName, String columnName, Boolean allowRetry) {
        return sharedSheetResponseService.getDeleteColumnWithIndexExtensionResponse(driveId, fileId, sheetName, columnName, allowRetry);
    }

    public ExtensionResponse getCompleteRowInfoExtensionResponse(String driveId, String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        return sharedSheetResponseService.getCompleteRowInfoExtensionResponse(driveId, fileId, sheetName, rowIndex, allowRetry);
    }

    public ExtensionResponse getDeleteRowWithIndexExtensionResponse(String driveId, String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        return sharedSheetResponseService.getDeleteRowWithIndexExtensionResponse(driveId, fileId, sheetName, rowIndex, allowRetry);
    }
}