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

public class SharedAreaSubCatalogRequests {

    private static final Logger LOGGER = LoggerFactory.getLogger(SharedAreaSubCatalogRequests.class);

    private final ReentryHandlerFactory reentryHandlerFactory;

    @Inject
    public SharedAreaSubCatalogRequests(ReentryHandlerFactory reentryHandlerFactory) {
        this.reentryHandlerFactory = reentryHandlerFactory;
    }

    @SubCatalogRequest(name = SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_TABLE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter table data and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddRowsToTable(
            @Field.Desc(name = "inputMap", type = "{ REENTER : Boolean, State Id : Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text, CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddRowsToTable start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleConfirmation(map);
    }

    @SubCatalogRequest(name = SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_TABLE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter table data")
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddRowsToTable(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text, Rows: [ Text ], CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddRowsToTable start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter sheet data")
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddRowsToSheet(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Rows: [ Text ], REENTER: Boolean, CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddRowsToSheet start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter sheet data")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddRowsToSheet(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Rows: [ Text ], CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddRowsToSheet start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_RANGE_SUM,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter range sum data")
    @Field(name = "Sum", type = "Number", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterRangeSum(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Range: Text, REENTER: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterRangeSum start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleConfirmation(map);
    }

    @SubCatalogRequest(name = SubCatalogConstants.HANDLE_REENTER_RANGE_SUM,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter range sum data")
    @Field(name = "Sum", type = "Number", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterRangeSum(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterRangeSum start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_TABLE_ROWS,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter table data for get table rows and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetTableRows(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, " +
                    "File Id: Text, Table Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetTableRows start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleGetTableRowsConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_TABLE_ROWS,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get table rows")
    @Field.Desc(name = "Table Rows", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetTableRows(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetTableRows start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleGetTableRowsReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_SHARED_SHEET_VLOOKUP,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for sheet vlookup and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterSheetVlookup(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, " +
                    "File Id: Text, Lookup Value: Text, Sheet Name: Text, " +
                    "Range: Text, Column Index: Number, Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterSheetVlookup start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleSheetVlookupConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_SHARED_SHEET_VLOOKUP,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter sheet vlookup")
    @Field.Desc(name = "Vlookup Result", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterSheetVlookup(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Lookup Value: Text, Sheet Name: Text, Range: Text, " +
                    "Column Index: Number, Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterSheetVlookup start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleSheetVlookupReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_DATA_TO_CELL,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add data to cell and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddDataToCell(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Cell Id: Text, Data: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddDataToCell start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddDataToCellConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_DATA_TO_CELL,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add data to cell")
    @Field.Desc(name = "Cell Information", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddDataToCell(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell Id: Text, Data: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddDataToCell start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddDataToCellReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for get rows for range and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetRowsForRange(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetRowsForRange start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleGetRowsForRangeConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_ROWS_FOR_RANGE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get rows for range")
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetRowsForRange(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetRowsForRange start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleGetRowsForRangeReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_VLOOKUP_IN_TABLE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for vlookup in table and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterVlookupInTable(
            @Field.Desc(name = "inputMap", type = "{REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Lookup Value: Text, " +
                    "Table Name: Text, Column Index: Number, Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterVlookupInTable start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleVlookupInTableConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_VLOOKUP_IN_TABLE,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter vlookup in table")
    @Field.Desc(name = "Vlookup Result", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterVlookupInTable(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Lookup Value: Text, Table Name: Text, Column Index: Number, " +
                    "Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterVlookupInTable start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleVlookupInTableReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_DELETE_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for delete row with index and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterDeleteRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterDeleteRowWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleDeleteRowWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_DELETE_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter delete row with index")
    @Field.Boolean(name = "Deleted", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterDeleteRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterDeleteRowWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleDeleteRowWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_COMPLETE_ROW_INFO,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for get complete row info and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetCompleteRowInfo(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetCompleteRowInfo start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleGetCompleteRowInfoConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_COMPLETE_ROW_INFO,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter get complete row info")
    @Field(name = "Row Data", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetCompleteRowInfo(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetCompleteRowInfo start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleGetCompleteRowInfoReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_DELETE_COLUMN,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for delete column and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterDeleteColumn(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Table Name: Text, " +
                    "Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterDeleteColumn start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleDeleteColumnConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_DELETE_COLUMN,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter delete column")
    @Field(name = "Is Deleted", type = "Switch", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterDeleteColumn(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text, Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterDeleteColumn start: {}", map);

        TableReentryHandler handler = reentryHandlerFactory.getTableReentryHandler();
        return handler.handleDeleteColumnReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_EXECUTE_FORMULA,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for execute formula and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterExecuteFormula(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Formula: Text, " +
                    "Formula Name: Text, Comment: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterExecuteFormula start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleExecuteFormulaConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_EXECUTE_FORMULA,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter execute formula")
    @Field(name = "Formula Result", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterExecuteFormula(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Formula: Text, Formula Name: Text, Comment: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterExecuteFormula start: {}", map);

        RangeReentryHandler handler = reentryHandlerFactory.getRangeReentryHandler();
        return handler.handleExecuteFormulaReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_SEARCH,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for search and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterSearch(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Data To Search: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterSearch start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleSearchConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_SEARCH,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter search")
    @Field.Desc(name = "Cell Ids", type = "[ Text ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterSearch(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Data To Search: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterSearch start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleSearchReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_CELL_DATA,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for cell data and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterCellData(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Cell Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterCellData start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleCellDataConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_CELL_DATA,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter cell data")
    @Field(name = "Cell Information", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterCellData(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterCellData start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleCellDataReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_SHARED_EXCEL_FILE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for get shared Excel file and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetSharedExcelFile(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Shared Folder Name: Text, Shared File Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetSharedExcelFile start: {}", map);

        SharedFileReentryHandler handler = reentryHandlerFactory.getSharedFileReentryHandler();
        return handler.handleGetSharedExcelFileConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_SHARED_EXCEL_FILE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter get shared Excel file")
    @Field.Desc(name = "Shared Excel File", type = "Entity(Excel File)", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetSharedExcelFile(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Shared Folder Name: Text, " +
                    "Shared File Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetSharedExcelFile start: {}", map);

        SharedFileReentryHandler handler = reentryHandlerFactory.getSharedFileReentryHandler();
        return handler.handleGetSharedExcelFileReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Confirm reenter get column values")
    @Field.Desc(name = "Result", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetColumnValues(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text, Key Rows: [ Text ], Lookup Column: [ Text ], " +
                    "Get All Record: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetColumnValues start: {}", map);

        ColumnReentryHandler handler = reentryHandlerFactory.getColumnReentryHandler();
        return handler.handleGetColumnValuesConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get column values")
    @Field.Desc(name = "Result", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetColumnValues(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text, Key Rows: [ Text ], Lookup Column: [ Text ], " +
                    "Get All Record: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetColumnValues start: {}", map);

        ColumnReentryHandler handler = reentryHandlerFactory.getColumnReentryHandler();
        return handler.handleGetColumnValuesReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_COLUMNS_SUM,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter columns sum")
    @Field.Desc(name = "Sum", type = "Number", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterColumnsSum(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text, Columns: [ Text ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterColumnsSum start: {}", map);

        ColumnReentryHandler handler = reentryHandlerFactory.getColumnReentryHandler();
        return handler.handleColumnsSumConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_COLUMNS_SUM,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter columns sum")
    @Field.Desc(name = "Sum", type = "Number", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterColumnsSum(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Table Name: Text, Columns: [ Text ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterColumnsSum start: {}", map);

        ColumnReentryHandler handler = reentryHandlerFactory.getColumnReentryHandler();
        return handler.handleColumnsSumReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_CONVERT_TO_PDF,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter convert to PDF")
    @Field.Desc(name = "PDF File", type = "File", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterConvertToPdf(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text" +
                    ", File Id: Text }") Map<String, Object> map) {

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
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterConvertToPdf start: {}", map);

        PdfReentryHandler handler = reentryHandlerFactory.getPdfReentryHandler();
        return handler.handleConvertToPdfReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_DELETE_COLUMN_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter delete column with index")
    @Field.Desc(name = "Deleted", type = "Boolean", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterDeleteColumnWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterDeleteColumnWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleDeleteColumnWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_DELETE_COLUMN_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter delete column with index")
    @Field.Desc(name = "Deleted", type = "Boolean", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterDeleteColumnWithIndex(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterDeleteColumnWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleDeleteColumnWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter add row with index data")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Rows: [ Text ], CSV File: File, " +
                    "Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddRowWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleAddRowWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add row with index data")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Rows: [ Text ], CSV File: File, " +
                    "Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddRowWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleAddRowWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_UPDATE_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter update row with index data")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterUpdateRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterUpdateRowWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleUpdateRowWithIndexConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_UPDATE_ROW_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter update row with index data")
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterUpdateRowWithIndex(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Row: Text, Row Index: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterUpdateRowWithIndex start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleUpdateRowWithIndexReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES_IN_SHEET,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Confirm reenter get column values in sheet")
    @Field.Desc(name = "Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetColumnValuesInSheet(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Range: Text, Lookup Rows: [ Text ], Lookup Columns: [ Text ], " +
                    "Get All Record: Boolean, Column Number: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetColumnValuesInSheet start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleGetColumnValuesInSheetConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES_IN_SHEET,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get column values in sheet")
    @Field.Desc(name = "Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetColumnValuesInSheet(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Range: Text, Lookup Rows: [ Text ], Lookup Columns: [ Text ], " +
                    "Get All Record: Boolean, Column Number: Number }")
            Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetColumnValuesInSheet start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleGetColumnValuesInSheetReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_MULTIPLE_CELL_DATA,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add multiple cell data and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddMultipleCellData(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell ID Data: [ { Cell Id: Text, Data: Text } ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddMultipleCellData start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddMultipleCellDataConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_MULTIPLE_CELL_DATA,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add multiple cell data")
    @Field.Desc(name = "Cell Data", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddMultipleCellData(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Cell ID Data: [ { Cell Id: Text, Data: Text } ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddMultipleCellData start: {}", map);

        CellReentryHandler handler = reentryHandlerFactory.getCellReentryHandler();
        return handler.handleAddMultipleCellDataReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_CLEAR_SHARED_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for clear shared sheet and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterClearSharedSheet(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, " +
                    "Drive Id: Text, File Id: Text, Sheet Name: Text, " +
                    "Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterClearSharedSheet start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleClearSharedSheetConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_CLEAR_SHARED_SHEET,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter clear shared sheet")
    @Field.Boolean(name = "Cleared", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterClearSharedSheet(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, File Id: Text, " +
                    "Sheet Name: Text, Range: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterClearSharedSheet start: {}", map);

        SheetReentryHandler handler = reentryHandlerFactory.getSheetReentryHandler();
        return handler.handleClearSharedSheetReentry(map);
    }
}
