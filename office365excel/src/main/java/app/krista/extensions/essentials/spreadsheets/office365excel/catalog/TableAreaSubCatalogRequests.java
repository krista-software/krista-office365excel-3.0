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

import java.util.Map;

public class TableAreaSubCatalogRequests {

    private static final Logger LOGGER = LoggerFactory.getLogger(TableAreaSubCatalogRequests.class);
    private final ReentryHandlerFactory reentryHandlerFactory;

    @Inject
    public TableAreaSubCatalogRequests(ReentryHandlerFactory reentryHandlerFactory) {
        this.reentryHandlerFactory = reentryHandlerFactory;
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_TABLE_VLOOKUP,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for table vlookup and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterTableVlookup(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, "
                    + "File Id: Text, Lookup Value: Text, Table Name: Text, "
                    + "Column Index: Number, Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterTableVlookup start: {}", map);

        TableVlookupReentryHandler handler = reentryHandlerFactory.getTableVlookupReentryHandler();
        return handler.handleTableVlookupConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_TABLE_VLOOKUP,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter table vlookup")
    @Field(name = "Lookup Result", type = "FreeForm", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterTableVlookup(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, "
                    + "Lookup Value: Text, Table Name: Text, "
                    + "Column Index: Number, Exact Match: Boolean }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterTableVlookup start: {}", map);

        TableVlookupReentryHandler handler = reentryHandlerFactory.getTableVlookupReentryHandler();
        return handler.handleTableVlookupReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_IN_TABLE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add rows and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddRows(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, "
                    + "File Id: Text, Table Name: Text, Rows: [ Text ], "
                    + "CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddRows start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleAddRowsConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_IN_TABLE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add rows")
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddRows(
            @Field.Desc(name = "inputMap", type = "{State Id: Text, File Id: Text, "
                    + "Table Name: Text, Rows: [ Text ], "
                    + "CSV File: File }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddRows start: {}", map);

        RowReentryHandler handler = reentryHandlerFactory.getRowReentryHandler();
        return handler.handleAddRowsReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Checks if user wants to re-enter data for add rows to table with index and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterAddRowsToTableWithIndex(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, Drive Id: Text, "
                    + "File Id: Text, Table Name: Text, Rows: [ Text ], "
                    + "CSV File: File, Row Index: Number, Separator: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterAddRowsToTableWithIndex start: {}", map);

        TableRowsReentryHandler handler = reentryHandlerFactory.getTableRowsReentryHandler();
        return handler.handleAddRowsToTableConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_ADD_ROWS_TO_TABLE_WITH_INDEX,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter add rows to table with index")
    @Field.Desc(name = "Total Rows", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterAddRowsToTableWithIndex(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, Drive Id: Text, "
                    + "File Id: Text, Table Name: Text, Rows: [ Text ], "
                    + "CSV File: File, Row Index: Number, Separator: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterAddRowsToTableWithIndex start: {}", map);

        TableRowsReentryHandler handler = reentryHandlerFactory.getTableRowsReentryHandler();
        return handler.handleAddRowsToTableReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_ROWS,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Checks if user wants to re-enter data for get rows and if yes, sends prompt to do so")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetRows(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, "
                    + "File Id: Text, Table Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetRows start: {}", map);

        TableRowsReentryHandler handler = reentryHandlerFactory.getTableRowsReentryHandler();
        return handler.handleConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_ROWS,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get rows")
    @Field.Desc(name = "Table Rows", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetRows(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, File Id: Text, "
                    + "Table Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetRows start: {}", map);

        TableRowsReentryHandler handler = reentryHandlerFactory.getTableRowsReentryHandler();
        return handler.handleReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_TABLE_SUMMATION,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter table summation data")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterTableSummation(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, "
                    + "File Id: Text, Table Name: Text, Columns: [ Text ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterTableSummation start: {}", map);

        SummationReentryHandler handler = reentryHandlerFactory.getSummationReentryHandler();
        return handler.handleTableSummationConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_TABLE_SUMMATION,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter table summation data")
    @Field(name = "Sum", type = "Number", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterTableSummation(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, "
                    + "File Id: Text, Table Name: Text, Columns: [ Text ] }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterTableSummation start: {}", map);

        SummationReentryHandler handler = reentryHandlerFactory.getSummationReentryHandler();
        return handler.handleTableSummationReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_DELETE_TABLE_COLUMN,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter delete table column data")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterDeleteTableColumn(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, "
                    + "File Id: Text, Table Name: Text, Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterDeleteTableColumn start: {}", map);

        DeleteColumnReentryHandler handler = reentryHandlerFactory.getTableColumnReentryHandler();
        return handler.handleDeleteTableColumnConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_DELETE_TABLE_COLUMN,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter delete table column data")
    @Field.Boolean(name = "Is Deleted", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterDeleteTableColumn(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, "
                    + "File Id: Text, Table Name: Text, Column Name: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterDeleteTableColumn start: {}", map);

        DeleteColumnReentryHandler handler = reentryHandlerFactory.getTableColumnReentryHandler();
        return handler.handleDeleteTableColumnReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_CREATE_TABLE_FOR_RANGE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Confirm reenter create table for range data")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterCreateTableForRange(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, "
                    + "File Id: Text, Sheet Name: Text, Range: Text, "
                    + "Drive Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterCreateTableForRange start: {}", map);

        TableCreationReentryHandler handler = reentryHandlerFactory.getTableCreationReentryHandler();
        return handler.handleCreateTableForRangeConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_CREATE_TABLE_FOR_RANGE,
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            description = "Handle reenter create table for range data")
    @Field.Desc(name = "Table", type = "Entity(Table)", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterCreateTableForRange(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, "
                    + "File Id: Text, Sheet Name: Text, Range: Text, "
                    + "Drive Id: Text }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterCreateTableForRange start: {}", map);

        TableCreationReentryHandler handler = reentryHandlerFactory.getTableCreationReentryHandler();
        return handler.handleCreateTableForRangeReentry(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.CONFIRM_REENTER_GET_COLUMN_VALUES_FOR_KEYS,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Confirm reenter get column values for keys data")
    @SuppressWarnings("unchecked")
    public ExtensionResponse confirmReenterGetColumnValuesForKeys(
            @Field.Desc(name = "inputMap", type = "{ REENTER: Boolean, State Id: Text, "
                    + "File Id: Text, Table Name: Text, Key Rows: [ Text ], "
                    + "Lookup Columns: [ Text ], Get All Record: Boolean, "
                    + "Column Number: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest confirmReenterGetColumnValuesForKeys start: {}", map);

        ColumnValuesReentryHandler handler = reentryHandlerFactory.getColumnValuesReentryHandler();
        return handler.handleGetColumnValuesForKeysConfirmation(map);
    }

    @SubCatalogRequest(
            name = SubCatalogConstants.HANDLE_REENTER_GET_COLUMN_VALUES_FOR_KEYS,
            type = CatalogRequest.Type.QUERY_SYSTEM,
            description = "Handle reenter get column values for keys data")
    @Field.Desc(name = "Result", type = "[ Composite ]", required = false)
    @SuppressWarnings("unchecked")
    public ExtensionResponse handleReenterGetColumnValuesForKeys(
            @Field.Desc(name = "inputMap", type = "{ State Id: Text, "
                    + "File Id: Text, Table Name: Text, Key Rows: [ Text ], "
                    + "Lookup Columns: [ Text ], Get All Record: Boolean, "
                    + "Column Number: Number }") Map<String, Object> map) {

        LOGGER.info("SubCatalogRequest handleReenterGetColumnValuesForKeys start: {}", map);

        ColumnValuesReentryHandler handler = reentryHandlerFactory.getColumnValuesReentryHandler();
        return handler.handleGetColumnValuesForKeysReentry(map);
    }
}
