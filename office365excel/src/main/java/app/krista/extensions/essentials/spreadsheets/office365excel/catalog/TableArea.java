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

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.ExtensionResponse;
import app.krista.extension.executor.ExtensionResponseBuilder;
import app.krista.extension.executor.Invoker;
import app.krista.extension.impl.anno.Attribute;
import app.krista.extension.impl.anno.CatalogRequest;
import app.krista.extension.impl.anno.Domain;
import app.krista.extension.impl.anno.Field;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TableOperationsFacade;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.util.EventHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import com.kristasoft.common.holders.ThreadLocalProxy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Domain(id = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24",
        name = "Spreadsheets",
        ecosystemId = "catEntryEcosystem_84b53163-327b-4b1b-8c96-9334d292f9f5",
        ecosystemName = "Essentials",
        ecosystemVersion = "4ab34479-fc9d-49f6-abda-4762af24ca28")
public class TableArea {

    private final WorkbookService workbookService;
    private static final Logger LOGGER = LoggerFactory.getLogger(TableArea.class);
    private final ExecutorService executorService;
    private final EventHandler eventHandler;
    private final TableOperationsFacade tableOperationsFacade;
    private final Office365ExcelAuditService auditService;
    private final String invokerId;
    private final RequestContext requestContext;
    private final AuthorizationContext authorizationContext;
    private final GraphServiceClientProviderFactory providerFactory;

    @Inject
    public TableArea(WorkbookService workbookService, EventHandler eventHandler, TableOperationsFacade tableOperationsFacade,
                     Office365ExcelAuditService auditService, RequestContext requestContext,
                     AuthorizationContext authorizationContext, GraphServiceClientProviderFactory providerFactory,
                     Invoker invoker) {
        this.workbookService = workbookService;
        this.eventHandler = eventHandler;
        this.tableOperationsFacade = tableOperationsFacade;
        this.executorService = Executors.newFixedThreadPool(10);
        this.auditService = auditService;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.providerFactory = providerFactory;
        this.invokerId = invoker.getInvokerId();
    }

    @CatalogRequest(
            id = "localDomainRequest_0cf8ac18-8046-4aea-a8fb-ada945c4ba9d",
            name = "Table Vlookup",
            description = "Table Vlookup",
            area = "Table",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field(name = "Lookup Result", type = "FreeForm", required = false, options = {})
    public ExtensionResponse tableVlookup(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Lookup Value", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String lookupValue,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field(name = "Column Index", type = "Number", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double columnIndex,
            @Field.Boolean(name = "Exact Match", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean exactMatch,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Table Vlookup", () -> {
            LOGGER.info("TableArea.tableVlookup start: fileID={}, lookupValue={}, tableName={}, columnIndex={}, exactMatch={}, allowRetry={}",
                    safeString(fileId), safeString(lookupValue), safeString(tableName), columnIndex, exactMatch, allowRetry);
            return this.tableOperationsFacade.getTableVlookupExtensionResponse(fileId, lookupValue, tableName, columnIndex, exactMatch, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_db9bbe56-1b7d-485c-80fd-edc99e83987e",
            name = "Add Rows",
            description = "Add Rows In Table",
            area = "Table",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse addRows(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.File(name = "Rows From Csv", multipleFileUpload = true, required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) List<File> rowsFromCsv,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Add Rows", () -> {
            LOGGER.info("TableArea.addRows start: fileID={}, tableName={}, allowRetry={}", safeString(fileID), safeString(tableName), allowRetry);
            return tableOperationsFacade.getAddRowsExtensionResponse(fileID, tableName, rows, rowsFromCsv, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_f23738e6-7a24-4e75-beb4-d98e71e35d16",
            name = "Add Rows Async",
            description = "Add Rows Async",
            area = "Table",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addRowsAsync(
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the table'")}) String fileId,
            @Field.Text(name = "Table Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the table to add rows to'")}) String tableName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.Desc(name = "Rows From Csv", type = "[ File ]", required = false) List<File> rowsFromCsv) {
        return audited("Add Rows Async", () -> {
            if (!Worksheets.isRowDataAvailable(rows, rowsFromCsv)) {
                throw new IllegalArgumentException("Row data not provided");
            }
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);

                    List<FreeForm> result = workbookService.addRowsInTable(null,
                            fileId, tableName, rows, rowsFromCsv, null, Constants.COMMA);

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Total Rows", "[ Composite ]", result);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (Exception cause) {
                    LOGGER.error("Error in addRowsAsync with fileId: {}, tableName: {}: {}",
                            fileId, tableName, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_db9bbe56-1b7d-485c-80fd-edc99e83987e_result",
            name = "Get Add Rows Result",
            description = "Get the result of the asynchronous Add Rows operation.",
            area = "Table",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field.Desc(name = "Total Rows", type = "[ Composite ]", required = false)
    public List<FreeForm> getAddRowsResult(
            @Field(name = "eventName", type = "Text", attributes = {@Attribute(name = "toolTip", value = "'The name of the event to wait for'")}) String eventName,
            @Field(name = "eventData", type = "FreeForm", attributes = {@Attribute(name = "toolTip", value = "'The data returned by the asynchronous operation'")}) FreeForm eventData,
            @Field(name = "taskID", type = "Text", attributes = {@Attribute(name = "toolTip", value = "'The task ID returned by the async operation'")}) String taskID) {

        return audited("Get Add Rows Result", () -> {
            if (eventName.equals(taskID)) {
                return (List<FreeForm>) eventData.get("Total Rows");
            }
            throw new IllegalStateException("Invalid task ID");
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_cc8e655d-ac6b-4e0d-a790-d8562be61925",
            name = "Add Rows To Table",
            description = "Add Rows To Table Using Index",
            area = "Table",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Total Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse addRowsToTable(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Drive ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveID,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.File(name = "Rows From Csv", multipleFileUpload = true, required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) List<File> rowsFromCsv,
            @Field(name = "Index", type = "Number", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double index,
            @Field.Text(name = "Separator", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String separator,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Add Rows To Table", () -> {
            LOGGER.info("TableArea.addRowsToTable start: fileID={}, driveID={}, tableName={}, index={}, allowRetry={}", safeString(fileID),
                    safeString(driveID), safeString(tableName), index, allowRetry);
            return tableOperationsFacade.getAddRowsToTableExtensionResponse(fileID, driveID, tableName, rows, rowsFromCsv, index, separator, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_a16fa33c-5588-4185-a740-18522c26f043",
            name = "Add Rows To Table Async",
            description = "Add Rows To Table Async",
            area = "Table",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "TaskId", required = false)
    public String addRowsToTableAsync(
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the table'")}) String fileId,
            @Field.Text(name = "Drive Id", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "Table Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the table to add rows to'")}) String tableName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.Desc(name = "Rows From Csv", type = "[ File ]", required = false) List<File> rowsFromCsv,
            @Field(name = "Row Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The row position where new rows will be inserted (1-based index)'")}) Double rowIndex,
            @Field.Text(name = "Separator", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The delimiter used to separate values in CSV data (default: comma)'")}) String separator) {
        return audited("Add Rows To Table Async", () -> {
            if (!Worksheets.isRowDataAvailable(rows, rowsFromCsv)) {
                throw new IllegalArgumentException("Row data not provided");
            }
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);

                    // Batch processing for large datasets
                    List<FreeForm> result = workbookService.addRowsInTable(driveId, fileId, tableName, rows, rowsFromCsv,
                            rowIndex.intValue() - 1, separator == null ? Constants.COMMA : separator);

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Total Rows", "[ Composite ]", result);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (Exception cause) {
                    LOGGER.error("Error in addRowsToTableAsync with driveId: {}, fileId: {}, tableName: {}, rowIndex: {}: {}",
                            driveId, fileId, tableName, rowIndex, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_cc8e655d-ac6b-4e0d-a790-d8562be61925_result",
            name = "Get Add Rows To Table Result",
            description = "Get the result of the asynchronous Add Rows To Table operation.",
            area = "Table",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field.Desc(name = "Total Rows", type = "[ Composite ]", required = false)
    public List<FreeForm> getAddRowsToTableResult(
            @Field(name = "eventName", type = "Text", attributes = {@Attribute(name = "toolTip", value = "'The name of the event to wait for'")}) String eventName,
            @Field(name = "eventData", type = "FreeForm", attributes = {@Attribute(name = "toolTip", value = "'The data returned by the asynchronous operation'")}) FreeForm eventData,
            @Field(name = "taskID", type = "Text", attributes = {@Attribute(name = "toolTip", value = "'The task ID returned by the async operation'")}) String taskID) {

        return audited("Get Add Rows To Table Result", () -> {
            if (eventName.equals(taskID)) {
                return (List<FreeForm>) eventData.get("Total Rows");
            }
            throw new IllegalStateException("Invalid task ID");
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_46024512-85f4-471c-a552-e7a6299f9952",
            name = "Get Rows",
            description = "Return rows from the given table.",
            area = "Table",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Table Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse getRows(
            @Field.Text(name = "Workbook ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String workbookID,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Rows", () -> {
            LOGGER.info("TableArea.getRows start: workbookID={}, tableName={}, allowRetry={}", safeString(workbookID), safeString(tableName), allowRetry);
            return tableOperationsFacade.getTableRowsExtensionResponse(workbookID, tableName, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_97f13a73-d007-4207-9e45-fd0599a58a86",
            name = "Summation",
            description = "Sum Operation will perform for a provided range within table",
            area = "Table",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field(name = "Sum", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse summation(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Desc(name = "Columns", type = "[ Text ]", required = true) List<String> columns,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Summation", () -> {
            LOGGER.info("TableArea.summation start: fileID={}, tableName={}, columns={}, allowRetry={}", safeString(fileID), safeString(tableName), columns, allowRetry);
            return tableOperationsFacade.getTableSummationExtensionResponse(fileID, tableName, columns, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_679623a1-e913-4462-bc44-fa0dce6b5b2e",
            name = "Delete Column",
            description = "Delete a particular column from the given table from the spreadsheet.",
            area = "Table",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Boolean(name = "Is Deleted", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse deleteColumn(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Text(name = "Column Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String columnName,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Delete Column", () -> {
            LOGGER.info("TableArea.deleteColumn start: fileID={}, tableName={}, columnName={}, allowRetry={}",
                    safeString(fileID), safeString(tableName), safeString(columnName), allowRetry);
            return tableOperationsFacade.getDeleteTableColumnExtensionResponse(fileID, tableName, columnName, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_153abfdc-b246-48a0-bc03-fde390fdedd7",
            name = "Create Table for range",
            description = "Create table from a given range.",
            area = "Table",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Table", type = "Entity(Table)", required = false)
    public ExtensionResponse createTableForRange(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Range", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String range,
            @Field.Text(name = "Drive Id", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Create Table for range", () -> {
            LOGGER.info("TableArea.createTableForRange start: fileId={}, sheetName={}, range={}, driveId={}, allowRetry={}",
                    safeString(fileId), safeString(sheetName), safeString(range), safeString(driveId), allowRetry);
            return this.tableOperationsFacade.getCreateTableForRangeExtensionResponse(fileId, sheetName, range, driveId, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_05ee821f-0a20-407e-aab6-89425fff9527",
            name = "Get Column Values for Keys",
            description = "Get column values for key rows and lookup columns.",
            area = "Table",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Desc(name = "Result", type = "[ Composite ]", required = false)
    public ExtensionResponse getColumnValuesForKeys(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Desc(name = "Key Rows", type = "[ Text ]", required = true) List<String> keyRows,
            @Field.Desc(name = "Lookup Columns", type = "[ Text ]", required = true) List<String> lookupColumns,
            @Field.Boolean(name = "Get All Record", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean getAllRecord,
            @Field(name = "Column Number", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double columnNumber,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Get Column Values for Keys", () -> {
            LOGGER.info("TableArea.getColumnValuesForKeys start: fileID={}, tableName={}, keyRows={}, lookupColumns={}, getAllRecord={}, columnNumber={}, allowRetry={}",
                    safeString(fileID), safeString(tableName), keyRows, lookupColumns, getAllRecord, columnNumber, allowRetry);
            return tableOperationsFacade.getColumnValuesForKeysExtensionResponse(fileID, tableName, keyRows, lookupColumns, getAllRecord, columnNumber, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_01655859-1099-47e4-9b43-859860f9444f",
            name = "Add Rows To Table In Batch Async",
            description = "Adds multiple rows of data to an Excel table in a single asynchronous batch operation, allowing efficient bulk insertion without blocking the client.",
            area = "Table",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Text(name = "TaskId", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addRowsToTableInBatchAsync(
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the table'")}) String fileId,
            @Field.Text(name = "Drive Id", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "Table Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the table to add rows to'")}) String tableName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.Desc(name = "Rows From Csv", type = "[ File ]", required = false) List<File> rowsFromCsv,
            @Field(name = "Row Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The row position where new rows will be inserted (1-based index)'")}) Double rowIndex,
            @Field.Text(name = "Separator", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The delimiter used to separate values in CSV data (default: comma)'")}) String separator) {
        return audited("Add Rows To Table In Batch Async", () -> {
            if (!Worksheets.isRowDataAvailable(rows, rowsFromCsv)) {
                throw new IllegalArgumentException("Row data not provided");
            }
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    LOGGER.info("Staring current task on thread: {} = {}", Thread.currentThread().getName(), Thread.currentThread().threadId());
                    ThreadLocalProxy.setAll(threadLocals);
                    RequestContext requestContext = ThreadLocalProxy.getThreadLocal(RequestContext.class).get();
                    // Batch processing for large datasets with progress tracking
                    workbookService.addRowsInBatches(driveId, fileId, tableName, rows, rowsFromCsv,
                            rowIndex.intValue() - 1, separator == null ? Constants.COMMA : separator, eventHandler, taskId, requestContext);

                } catch (Exception e) {
                    LOGGER.error("Error in addRowsToTableInBatchAsync: {}", e.getMessage(), e);
                    throw new RuntimeException(e);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_53e488b6-b8cc-4d07-949d-49fcac8769aa",
            name = "Add Rows Async To Table Progress Update",
            description = "Triggers for when an update on adding rows to an excel table is available when using add rows async. Approximately every 20% of total rows to add.",
            area = "Table",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field(name = "Message", type = "Paragraph", required = false, attributes = {@Attribute(name = "visualWidth", value = "L")})
    @Field(name = "Current Index", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The intex of the last record added. '")})
    @Field(name = "Progress", type = "Percentage", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    @Field.Boolean(name = "Completed", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public ExtensionResponse addRowsAsyncToTableProgressUpdate(
            @Field(name = "eventName", type = "Text") String eventName,
            @Field(name = "eventData", type = "FreeForm") FreeForm eventData,
            @Field(name = "Task Id", type = "Identifier", attributes = {@Attribute(name = "visualWidth", value = "S")}) String taskId) {
        return audited("Add Rows Async To Table Progress Update", () -> {
            if (eventName.equals(taskId)) {
                return new ExtensionResponseBuilder().success((Map<String, Object>) eventData.get("Data")).build();
            }
            throw new IllegalStateException("Invalid task ID");
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_c64b82d4-da6c-416c-9589-39587f9b758d",
            name = "Add Rows Async To Table  Result Available",
            description = "Request Description",
            area = "Table",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field(name = "Total Rows", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'Total rows in the table'")})
    @Field(name = "Rows Added", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'number of rows added to the table'")})
    @Field.Boolean(name = "Success", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    @Field(name = "Message", type = "Paragraph", required = false, attributes = {@Attribute(name = "visualWidth", value = "L"), @Attribute(name = "toolTip", value = "'status message to explain success indicator'")})
    @Field(name = "Start Time", type = "Time", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    @Field(name = "End Time", type = "Time", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    @Field(name = "Duration", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'Duration of operation in Ms'")})
    @Field.Text(name = "Duration Text", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'Duration of operation in human text'")})
    public ExtensionResponse addRowsAsyncToTableResultAvailable(
            @Field(name = "eventName", type = "Text") String eventName,
            @Field(name = "eventData", type = "FreeForm") FreeForm eventData,
            @Field(name = "Task Id", type = "Identifier", attributes = {@Attribute(name = "visualWidth", value = "S")}) String taskId) {
        return audited("Add Rows Async To Table  Result Available", () -> {
            // Check if this is a completion event (eventName ends with "_completion")
            if (eventName.equals(taskId + "_completion")) {
                return new ExtensionResponseBuilder().success((Map<String, Object>) eventData.get("Data")).build();
            }
            throw new IllegalStateException("Invalid task ID or event type");
        });
    }

    @FunctionalInterface
    private interface ThrowingSupplier<T> {
        T get() throws Exception;
    }

    private <T> T audited(String catalogName, ThrowingSupplier<T> action) {
        LOGGER.info("CatalogRequest [{}] triggered", catalogName);
        auditService.startRequest(invokerId, catalogName);
        try {
            T result = action.get();
            if (result instanceof ExtensionResponse er && er.getResult() == ExtensionResponse.Result.FAILURE) {
                LOGGER.error("CatalogRequest [{}] failure", catalogName);
                auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", catalogName, null, null, null);
            } else {
                LOGGER.info("CatalogRequest [{}] success", catalogName);
                auditService.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", catalogName, null, null, null);
            }
            return result;
        } catch (MustAuthorizeException e) {
            LOGGER.error("CatalogRequest [{}] failure - {}: {}", catalogName, "MustAuthorizeException", e.getMessage());
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", catalogName,
                    e.getMessage(), "MustAuthorizeException", null);
            throw e;
        } catch (Exception e) {
            LOGGER.error("CatalogRequest [{}] failure - {}: {}", catalogName, e.getClass().getSimpleName(), e.getMessage());
            auditService.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", catalogName,
                    e.getMessage(), e.getClass().getSimpleName(), null);
            if (e instanceof RuntimeException re) throw re;
            throw new RuntimeException(e);
        }
    }

}
