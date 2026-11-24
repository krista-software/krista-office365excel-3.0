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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sheetoperations.*;
import app.krista.model.base.File;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.util.List;
import java.util.Map;

@Service
public class SheetOperationsFacade {

    private final CellResponseService cellResponseService;
    private final WorksheetResponseService worksheetResponseService;
    private final RowResponseService rowResponseService;
    private final ColumnResponseService columnResponseService;
    private final FormulaResponseService formulaResponseService;
    private final SearchResponseService searchResponseService;
    private final WorkbookResponseService workbookResponseService;
    private final PdfResponseService pdfResponseService;

    @Inject
    public SheetOperationsFacade(CellResponseService cellResponseService,
                                 WorksheetResponseService worksheetResponseService,
                                 RowResponseService rowResponseService,
                                 ColumnResponseService columnResponseService,
                                 FormulaResponseService formulaResponseService,
                                 SearchResponseService searchResponseService,
                                 WorkbookResponseService workbookResponseService,
                                 PdfResponseService pdfResponseService) {
        this.cellResponseService = cellResponseService;
        this.worksheetResponseService = worksheetResponseService;
        this.rowResponseService = rowResponseService;
        this.columnResponseService = columnResponseService;
        this.formulaResponseService = formulaResponseService;
        this.searchResponseService = searchResponseService;
        this.workbookResponseService = workbookResponseService;
        this.pdfResponseService = pdfResponseService;
    }

    // Cell Operations
    public ExtensionResponse getGetCellValueExtensionResponse(String fileId, String sheetName, String cellId, Boolean allowRetry) {
        return cellResponseService.getGetCellValueExtensionResponse(fileId, sheetName, cellId, allowRetry);
    }

    public ExtensionResponse getAddDataToCellExtensionResponse(String fileID, String sheetName, String cellID, String data, Boolean allowRetry) {
        return cellResponseService.getAddDataToCellExtensionResponse(fileID, sheetName, cellID, data, allowRetry);
    }

    public ExtensionResponse getAddMultipleCellDataExtensionResponse(String fileId, String sheetName, List<Map<String, Object>> cellIDData, Boolean allowRetry) {
        return cellResponseService.getAddMultipleCellDataExtensionResponse(fileId, sheetName, cellIDData, allowRetry);
    }

    public ExtensionResponse getAddFormulaToCellExtensionResponse(String fileID, String sheetName, String cellID, String formula, Boolean allowRetry) {
        return cellResponseService.getAddFormulaToCellExtensionResponse(fileID, sheetName, cellID, formula, allowRetry);
    }

    // Worksheet Operations
    public ExtensionResponse getDeleteSheetExtensionResponse(String fileId, String sheetName, Boolean allowRetry) {
        return worksheetResponseService.getDeleteSheetExtensionResponse(fileId, sheetName, allowRetry);
    }

    public ExtensionResponse getCreateSheetExtensionResponse(String fileId, String newSheetName, Boolean allowRetry) {
        return worksheetResponseService.getCreateSheetExtensionResponse(fileId, newSheetName, allowRetry);
    }

    public ExtensionResponse getGetSheetsInformationExtensionResponse(String fileId, Boolean allowRetry) {
        return worksheetResponseService.getGetSheetsInformationExtensionResponse(fileId, allowRetry);
    }

    public ExtensionResponse getClearSheetForRangeExtensionResponse(String fileID, String sheetName, String cellRange, Boolean allowRetry) {
        return worksheetResponseService.getClearSheetForRangeExtensionResponse(fileID, sheetName, cellRange, allowRetry);
    }

    // Row Operations
    public ExtensionResponse getAddRowsExtensionResponse(String fileId, String sheetName, List<String> rows, List<File> rowsFromCsv, Boolean allowRetry) {
        return rowResponseService.getAddRowsExtensionResponse(fileId, sheetName, rows, rowsFromCsv, allowRetry);
    }

    public ExtensionResponse getAddRowWithIndexExtensionResponse(String fileId, String sheetName, List<String> rows,
                                                                 List<File> rowsFromCsv, Double rowIndex, Boolean allowRetry) {
        return rowResponseService.getAddRowWithIndexExtensionResponse(fileId, sheetName, rows, rowsFromCsv, rowIndex, allowRetry);
    }

    public ExtensionResponse getGetCompleteRowInfoExtensionResponse(String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        return rowResponseService.getGetCompleteRowInfoExtensionResponse(fileId, sheetName, rowIndex, allowRetry);
    }

    public ExtensionResponse getGetRowsForRangeExtensionResponse(String fileID, String sheetName, String range, Boolean allowRetry) {
        return rowResponseService.getGetRowsForRangeExtensionResponse(fileID, sheetName, range, allowRetry);
    }

    public ExtensionResponse getGetRowsForRangeExtensionResponse(String driveId, String fileID, String sheetName, String range, Boolean allowRetry) {
        return rowResponseService.getGetRowsForRangeExtensionResponse(driveId, fileID, sheetName, range, allowRetry);
    }

    public ExtensionResponse getUpdateRowWithIndexExtensionResponse(String fileId, String sheetName, String row, Double rowIndex, Boolean allowRetry) {
        return rowResponseService.getUpdateRowWithIndexExtensionResponse(fileId, sheetName, row, rowIndex, allowRetry);
    }

    public ExtensionResponse getDeleteRowWithIndexExtensionResponse(String fileId, String sheetName, Double rowIndex, Boolean allowRetry) {
        return rowResponseService.getDeleteRowWithIndexExtensionResponse(fileId, sheetName, rowIndex, allowRetry);
    }

    // Column Operations
    public ExtensionResponse getDeleteColumnWithIndexExtensionResponse(String fileId, String sheetName, String columnName, Boolean allowRetry) {
        return columnResponseService.getDeleteColumnWithIndexExtensionResponse(fileId, sheetName, columnName, allowRetry);
    }

    public ExtensionResponse getGetColumnValuesForKeysInSheetExtensionResponse(String fileId, String sheetName, String range,
                                                                               List<String> keyRows, List<String> keyColumns,
                                                                               Boolean getAllRecords, Double columnNumber, Boolean allowRetry) {
        return columnResponseService.getGetColumnValuesForKeysInSheetExtensionResponse(fileId, sheetName, range, keyRows, keyColumns, getAllRecords, columnNumber, allowRetry);
    }

    // Formula Operations
    public ExtensionResponse getExecuteFormulaExtensionResponse(String fileId, String formula, String formulaName, String comment, Boolean allowRetry) {
        return formulaResponseService.getExecuteFormulaExtensionResponse(fileId, formula, formulaName, comment, allowRetry);
    }

    public ExtensionResponse getVlookupExtensionResponse(String fileID, String lookupValue, String sheetName, String range,
                                                         Double columnIndex, Boolean exactMatch, Boolean allowRetry) {
        return formulaResponseService.getVlookupExtensionResponse(fileID, lookupValue, sheetName, range, columnIndex, exactMatch, allowRetry);
    }

    public ExtensionResponse getSummationExtensionResponse(String fileID, String sheetName, String range, Boolean allowRetry) {
        return formulaResponseService.getSummationExtensionResponse(fileID, sheetName, range, allowRetry);
    }

    // Search Operations
    public ExtensionResponse getFindValueExtensionResponse(String fileId, String sheetName, String lookupValue, Boolean allowRetry) {
        return searchResponseService.getFindValueExtensionResponse(fileId, sheetName, lookupValue, allowRetry);
    }

    public ExtensionResponse getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(String fileId, String sheetName,
                                                                                        String keyColumnName, String columnRange, Boolean allowRetry) {
        return searchResponseService.getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(fileId, sheetName, keyColumnName, columnRange, allowRetry);
    }

    // Workbook Operations
    public ExtensionResponse getGetExcelFileInformationExtensionResponse(String path, String name, Boolean allowRetry) {
        return workbookResponseService.getGetExcelFileInformationExtensionResponse(path, name, allowRetry);
    }

    // PDF Operations
    public ExtensionResponse getUploadFileAndConvertToPDFExtensionResponse(File excelFile, String filePath, Boolean allowRetry) {
        return pdfResponseService.getUploadFileAndConvertToPDFExtensionResponse(excelFile, filePath, allowRetry);
    }

    public ExtensionResponse getConvertToPDFExtensionResponse(String fileId, Boolean allowRetry) {
        return pdfResponseService.getConvertToPDFExtensionResponse(fileId, allowRetry);
    }

}