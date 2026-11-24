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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extension.impl.anno.CatalogRequest;
import app.krista.extension.impl.anno.Field;
import app.krista.extension.impl.anno.SubCatalogRequest;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.ReentryHandlerFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.*;

public class SheetAreaSubCatalogRequests {

    private static final Logger LOGGER = LoggerFactory.getLogger(SheetAreaSubCatalogRequests.class);
    private final ReentryHandlerFactory reentryHandlerFactory;

    @Inject
    public SheetAreaSubCatalogRequests(ReentryHandlerFactory reentryHandlerFactory) {
        this.reentryHandlerFactory = reentryHandlerFactory;
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for delete sheet and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterDeleteSheet(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text," +
                    "File Id: Text, Sheet Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterDeleteSheet start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleDeleteSheetConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter delete sheet")
    @Field(name = "Result", type = "Text", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterDeleteSheet(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, " +
                    "File Id: Text, Sheet Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterDeleteSheet start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleDeleteSheetReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_CELL_VALUE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for get cell value and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetCellValue(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, " +
                    "Cell Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetCellValue start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleGetCellValueConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_CELL_VALUE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get cell value")
    @Field(name = "Cell Information", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetCellValue(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetCellValue start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleGetCellValueReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add rows and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddRows(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, " +
                    "Rows: [ Text ], CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddRows start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleAddRowsConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_ROWS,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add rows")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddRows(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Rows: [ Text ], " +
                    "CSV File: File}") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddRows start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleAddRowsReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_SHEET_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add row with index and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Row Index: Number, " +
                    "Rows: [ Text ], CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddRowWithIndex start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleAddRowWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_SHEET_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add row with index")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row Index: Number, Rows: [ Text ], " +
                    "CSV File: [ File ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddRowWithIndex start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleAddRowWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_UPDATE_SHEET_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for update row with index and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterUpdateRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Row: Text, " +
                    "Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterUpdateRowWithIndex start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleUpdateRowWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_UPDATE_SHEET_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter update row with index")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterUpdateRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterUpdateRowWithIndex start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleUpdateRowWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_COMPLETE_SHEET_ROW_INFO,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for get complete row info and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetCompleteRowInfo(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, " +
                    "Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetCompleteRowInfo start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleGetCompleteRowInfoConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_COMPLETE_SHEET_ROW_INFO,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter get complete row info")
    @Field(name = "Row Data", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetCompleteRowInfo(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetCompleteRowInfo start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleGetCompleteRowInfoReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for delete column with index and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterDeleteColumnWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterDeleteColumnWithIndex start: {}", map);

        ColumnReentryHandler handler = reentryHandlerFactory.getColumnReentryHandler();
        return handler.handleDeleteColumnWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET_COLUMN_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter delete column with index")
    @Field.Boolean(name = "Deleted", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterDeleteColumnWithIndex(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterDeleteColumnWithIndex start: {}", map);

        ColumnReentryHandler handler = reentryHandlerFactory.getColumnReentryHandler();
        return handler.handleDeleteColumnWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for delete row with index and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterDeleteRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterDeleteRowWithIndex start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleDeleteRowWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_DELETE_SHEET_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter delete row with index")
    @Field.Boolean(name = "Deleted", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterDeleteRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterDeleteRowWithIndex start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleDeleteRowWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_FIND_VALUE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Confirm reenter find value")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterFindValue(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Lookup Value: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterFindValue start: {}", map);

        FindValueReentryHandler handler = reentryHandlerFactory.getFindValueReentryHandler();
        return handler.handleConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_FIND_VALUE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter find value")
    @Field.Desc(name = "Cell IDs", type = "[ Text ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterFindValue(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Lookup Value: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterFindValue start: {}", map);

        FindValueReentryHandler handler = reentryHandlerFactory.getFindValueReentryHandler();
        return handler.handleReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_EXCEL_FILE_INFO,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Confirm reenter get Excel file information")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetExcelFileInfo(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Path: Text, Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetExcelFileInfo start: {}", map);

        WorkbookReentryHandler handler = reentryHandlerFactory.getWorkbookReentryHandler();
        return handler.handleGetExcelFileInfoConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_EXCEL_FILE_INFO,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get Excel file information")
    @Field.Desc(name = "Excel File", type = "Entity(Excel File)", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetExcelFileInfo(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Path: Text, Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetExcelFileInfo start: {}", map);

        WorkbookReentryHandler handler = reentryHandlerFactory.getWorkbookReentryHandler();
        return handler.handleGetExcelFileInfoReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_SHEETS_INFO,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for get sheets information and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetSheetsInfo(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetSheetsInfo start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleGetSheetsInfoConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_SHEETS_INFO,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get sheets information")
    @Field.Desc(name = "Sheets", type = "[ Entity(Excel Sheet) ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetSheetsInfo(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetSheetsInfo start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleGetSheetsInfoReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_RETRY_ADD_DATA_TO_CELL,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to retry add data to cell and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmRetryAddDataToCell(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Cell Id: Text, Data: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmRetryAddDataToCell start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddDataToCellRetryConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_RETRY_ADD_DATA_TO_CELL,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle retry add data to cell")
    @Field.Desc(name = "Cell Information", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleRetryAddDataToCell(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell Id: Text, Data: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleRetryAddDataToCell start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddDataToCellRetry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_RETRY_EXECUTE_FORMULA,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to retry execute formula and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmRetryExecuteFormula(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Formula: Text, Formula Name: Text, Comment: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmRetryExecuteFormula start: {}", map);

        FormulaReentryHandler handler = reentryHandlerFactory.getFormulaReentryHandler();
        return handler.handleExecuteFormulaRetryConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_RETRY_EXECUTE_FORMULA,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle retry execute formula")
    @Field.Desc(name = "Formula Result", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleRetryExecuteFormula(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, " +
                    "Formula: Text, Formula Name: Text, Comment: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleRetryExecuteFormula start: {}", map);

        FormulaReentryHandler handler = reentryHandlerFactory.getFormulaReentryHandler();
        return handler.handleExecuteFormulaRetry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_CREATE_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for create sheet and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterCreateSheet(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text," +
                    "File Id: Text, Sheet Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterCreateSheet start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleCreateSheetConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_CREATE_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter create sheet")
    @Field.Desc(name = "Sheet", type = "Entity(Excel Sheet)", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterCreateSheet(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, " +
                    "File Id: Text, Sheet Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterCreateSheet start: {}", map);

        WorksheetReentryHandler handler = reentryHandlerFactory.getWorksheetReentryHandler();
        return handler.handleCreateSheetReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_SHEET_VLOOKUP,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for sheet vlookup and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterSheetVlookup(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, File Id: Text, Lookup Value: Text, " +
                    "Sheet Name: Text, Range: Text, Column Index: Number, Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterSheetVlookup start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleSheetVlookupConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_SHEET_VLOOKUP,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter sheet vlookup")
    @Field.Desc(name = "Lookup Result", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterSheetVlookup(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, Lookup Value: Text, " +
                    "Sheet Name: Text, Range: Text, Column Index: Number, Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterSheetVlookup start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleSheetVlookupReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_CLEAR_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for clear sheet range and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterClearSharedSheet(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterClearSharedSheet start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleClearSharedSheetConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_CLEAR_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter clear sheet range")
    @Field.Boolean(name = "Cleared", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterClearSharedSheet(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterClearSharedSheet start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleClearSharedSheetReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for get rows for range and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetRowsForRange(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetRowsForRange start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleGetRowsForRangeConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter get rows for range")
    @Field.Desc(name = "Rows Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetRowsForRange(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetRowsForRange start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleGetRowsForRangeReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_FORMULA_TO_CELL,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add formula to cell and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddFormulaToCell(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell Id: Text, Formula: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddFormulaToCell start: {}", map);

        FormulaReentryHandler handler = reentryHandlerFactory.getFormulaReentryHandler();
        return handler.handleAddFormulaToCellConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_FORMULA_TO_CELL,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add formula to cell")
    @Field(name = "Cell Information", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddFormulaToCell(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Cell Id: Text, Formula: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddFormulaToCell start: {}", map);

        FormulaReentryHandler handler = reentryHandlerFactory.getFormulaReentryHandler();
        return handler.handleAddFormulaToCellReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_SUMMATION,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter summation data")
    @Field(name = "Sum", type = "Number", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterSummation(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterSummation start: {}", map);

        SummationReentryHandler handler = reentryHandlerFactory.getSummationReentryHandler();
        return handler.handleSummationConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_SUMMATION,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter summation data")
    @Field(name = "Sum", type = "Number", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterSummation(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterSummation start: {}", map);

        SummationReentryHandler handler = reentryHandlerFactory.getSummationReentryHandler();
        return handler.handleSummationReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_CONVERT_TO_PDF,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter convert to PDF")
    @Field.Desc(name = "PDF File", type = "File", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterConvertToPdf(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, File Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterConvertToPdf start: {}", map);

        PdfReentryHandler handler = reentryHandlerFactory.getPdfReentryHandler();
        return handler.handleConvertToPdfConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_CONVERT_TO_PDF,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter convert to PDF")
    @Field.Desc(name = "PDF File", type = "File", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterConvertToPdf(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterConvertToPdf start: {}", map);

        PdfReentryHandler handler = reentryHandlerFactory.getPdfReentryHandler();
        return handler.handleConvertToPdfReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter get column values for keys in sheet")
    @Field.Desc(name = "Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetColumnValuesForKeysInSheet(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Range: Text, " +
                    "Lookup Rows: [ Text ], Lookup Columns: [ Text ], " +
                    "Get All Record: Boolean, Column Number: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetColumnValuesForKeysInSheet start: {}", map);

        ColumnValuesReentryHandler handler = reentryHandlerFactory.getColumnValuesReentryHandler();
        return handler.handleGetColumnValuesForKeysInSheetConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_SHEET_COLUMN_VALUES_IN_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter get column values for keys in sheet")
    @Field.Desc(name = "Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetColumnValuesForKeysInSheet(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, Range: Text, " +
                    "Lookup Rows: [ Text ], Lookup Columns: [ Text ], " +
                    "Get All Record: Boolean, Column Number: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetColumnValuesForKeysInSheet start: {}", map);

        ColumnValuesReentryHandler handler = reentryHandlerFactory.getColumnValuesReentryHandler();
        return handler.handleGetColumnValuesForKeysInSheetReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter upload file and convert to PDF")
    @Field.File(name = "PDF File", required = false)
    @Field(name = "File Id", type = "Text", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterUploadFileAndConvertToPDF(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Excel File: File, File Path: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterUploadFileAndConvertToPDF start: {}", map);

        FileReentryHandler handler = reentryHandlerFactory.getFileReentryHandler();
        return handler.handleUploadFileAndConvertToPDFConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_UPLOAD_FILE_AND_CONVERT_TO_PDF,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter upload file and convert to PDF")
    @Field.File(name = "PDF File", required = false)
    @Field(name = "File Id", type = "Text", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterUploadFileAndConvertToPDF(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Excel File : File, " +
                    "File Path: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterUploadFileAndConvertToPDF start: {}", map);

        FileReentryHandler handler = reentryHandlerFactory.getFileReentryHandler();
        return handler.handleUploadFileAndConvertToPDFReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_MIN_MAX_VALUE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for get min-max value and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetMinMaxValue(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Key Column Name: Text, Column Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetMinMaxValue start: {}", map);

        MinMaxValueReentryHandler handler = reentryHandlerFactory.getMinMaxValueReentryHandler();
        return handler.handleConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_MIN_MAX_VALUE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get min-max value")
    @Field(name = "Result", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetMinMaxValue(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Key Column Name: Text, Column Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetMinMaxValue start: {}", map);

        MinMaxValueReentryHandler handler = reentryHandlerFactory.getMinMaxValueReentryHandler();
        return handler.handleReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add multiple cell data and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddMultipleCellData(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "File Id: Text, Sheet Name: Text, " +
                    "Cell ID Data: [ { Cell Id: Text, Data: Text } ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddMultipleCellData start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddMultipleCellDataConfirmationForSheet(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA_FOR_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add multiple cell data")
    @Field.Desc(name = "Cell Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddMultipleCellData(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell ID Data: [ { Cell Id: Text, Data: Text } ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddMultipleCellData start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddMultipleCellDataReentryForSheet(map);
    }
}
