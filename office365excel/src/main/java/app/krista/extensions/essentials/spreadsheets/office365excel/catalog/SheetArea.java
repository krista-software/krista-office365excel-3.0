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
import app.krista.extension.executor.Invoker;
import app.krista.extension.impl.anno.Attribute;
import app.krista.extension.impl.anno.CatalogRequest;
import app.krista.extension.impl.anno.Domain;
import app.krista.extension.impl.anno.Field;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.SheetOperationsFacade;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.DriveService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.SheetFormattingService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetOperation;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.util.EventHandler;
import app.krista.ksdk.entities.Entities;
import app.krista.model.base.EntityValue;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import app.krista.model.entity.EntityDefinition;
import com.kristasoft.common.holders.ThreadLocalProxy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Domain(id = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24",
        name = "Spreadsheets",
        ecosystemId = "catEntryEcosystem_84b53163-327b-4b1b-8c96-9334d292f9f5",
        ecosystemName = "Essentials",
        ecosystemVersion = "6ff1d0a5-77ce-4859-a39b-baa7a716f556")
public class SheetArea {

    private static final Logger LOGGER = LoggerFactory.getLogger(SheetArea.class);
    private final WorkbookService workbookService;
    private final WorksheetService worksheetService;
    private final Entities registry;
    private final RestClientProvider restClientProvider;
    private final WorksheetOperation worksheetOperation;
    private final StringUtil stringUtil;
    ExecutorService executorService = Executors.newSingleThreadExecutor();
    private final EventHandler eventHandler;
    private final SheetFormattingService sheetFormattingService;
    private final SheetOperationsFacade sheetOperationsFacade;
    private final DriveService driveService;
    private final Office365ExcelAuditService auditService;
    private final String invokerId;
    private final RequestContext requestContext;
    private final AuthorizationContext authorizationContext;
    private final GraphServiceClientProviderFactory providerFactory;

    @Inject
    public SheetArea(Entities entities, WorkbookService workbookService, WorksheetService worksheetService,
                     RestClientProvider restClientProvider, WorksheetOperation worksheetOperation, StringUtil stringUtil,
                     EventHandler eventHandler, SheetFormattingService sheetFormattingService,
                     SheetOperationsFacade sheetOperationsFacade, DriveService driveService,
                     Office365ExcelAuditService auditService,
                     RequestContext requestContext, AuthorizationContext authorizationContext,
                     GraphServiceClientProviderFactory providerFactory, Invoker invoker) {
        this.workbookService = workbookService;
        this.worksheetService = worksheetService;
        this.registry = entities;
        this.restClientProvider = restClientProvider;
        this.worksheetOperation = worksheetOperation;
        this.stringUtil = stringUtil;
        this.eventHandler = eventHandler;
        this.sheetFormattingService = sheetFormattingService;
        this.sheetOperationsFacade = sheetOperationsFacade;
        this.driveService = driveService;
        this.auditService = auditService;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.providerFactory = providerFactory;
        this.invokerId = invoker.getInvokerId();
    }

    @CatalogRequest(
            id = "localDomainRequest_f2c9ef0f-467f-49e5-8b03-5110f244c6ac",
            name = "Get All Excel Files For Name",
            description = "Retrieve all available workbooks from your OneDrive based on name provided in input.",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Files", type = "[ Entity(Excel File) ]", required = false)
    public List<ExcelFile> getAllExcelFilesForName(
            @Field(name = "Name", type = "Text") String name) {
        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get All Excel Files For Name", () -> workbookService.getWorkbooks(name));
    }

    @CatalogRequest(
            id = "localDomainRequest_ceaa8b08-476b-4ec4-9e7c-8c07907f008b",
            name = "Delete Sheet",
            description = "Delete Sheet from File",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Text(name = "Result", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse deleteSheet(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Delete Sheet", () -> {
            LOGGER.info("SheetArea.deleteSheet start: fileId={}, sheetName={}, allowRetry={}", safeString(fileId), safeString(sheetName), allowRetry);
            return sheetOperationsFacade.getDeleteSheetExtensionResponse(fileId, sheetName, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_d425d8df-c949-44cf-9501-b9e4fc6be92d",
            name = "Get Cell Value",
            description = "Get Cell Value",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field(name = "Cell Information", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse getCellValue(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Cell ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String cellID,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Get Cell Value", () -> {
            LOGGER.info("SheetArea.getCellValue start: fileId={}, sheetName={}, cellID={}, allowRetry={}", safeString(fileID), safeString(sheetName), safeString(cellID), allowRetry);
            return sheetOperationsFacade.getGetCellValueExtensionResponse(fileID, sheetName, cellID, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_ccbb6553-f1c0-4773-b344-66b85593961f",
            name = "Add Rows",
            description = "Add Rows",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    public ExtensionResponse addRows(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.File(name = "Rows from Csv", multipleFileUpload = true, required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) List<File> rowsFromCsv,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Add Rows", () -> {
            LOGGER.info("SheetArea.addRows start: fileId={}, sheetName={}, allowRetry={}", safeString(fileID), safeString(sheetName), allowRetry);
            return sheetOperationsFacade.getAddRowsExtensionResponse(fileID, sheetName, rows, rowsFromCsv, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_61f9b0de-68f2-4ca3-82f6-10c18a0910de",
            name = "Add Rows Async",
            description = "Add Rows Async",
            area = "Sheet",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addRowsAsync(
            @Field.Text(name = "File ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileID,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String sheetName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.Desc(name = "Rows from Csv", type = "[ File ]", required = false) List<File> rowsFromCsv) {
        return audited("Add Rows Async", () -> {
            if (!Worksheets.isRowDataAvailable(rows, rowsFromCsv)) {
                throw new IllegalArgumentException(Constants.ROW_DATA_NOT_PROVIDED);
            }
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);

                    List<FreeForm> result = worksheetService.addRows(null, fileID, sheetName, rows, rowsFromCsv, null);
                    System.out.println("Result for the add rows " + result);
                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Result", "[ Composite ]", result);
                    System.out.println("Freeform " + freeForm);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (IOException cause) {
                    LOGGER.error("Error in addRowsAsync with fileID: {}, sheetName: {}: {}",
                            fileID, sheetName, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_ccbb6553-f1c0-4773-b344-66b85593961f_result",
            name = "Get Result",
            description = "Get the result of the asynchronous.",
            area = "Sheet",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field.Desc(name = "Total Rows", type = "[ Composite ]", required = false)
    public List<FreeForm> getAddRowsResult(
            @Field(name = "eventName", type = "Text") String eventName,
            @Field(name = "eventData", type = "FreeForm") FreeForm eventData,
            @Field(name = "taskID", type = "Text") String taskID) {
        return audited("Get Result", () -> {
            if (eventName.equals(taskID)) {
                return (List<FreeForm>) eventData.get("Result");
            }
            throw new IllegalStateException("Task ID mismatch. The provided task ID does not match the expected event.");
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_735f74ed-bc25-4eac-a401-7be10b880a60",
            name = "Add Row With Index",
            description = "Insert Row At The Desired Index",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    public ExtensionResponse addRowWithIndex(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field(name = "Row Index", type = "Number", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double rowIndex,
            @Field.File(name = "Rows From Csv", multipleFileUpload = true, required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) List<File> rowsFromCsv,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Add Row With Index", () -> {
            LOGGER.info("SheetArea.addRowWithIndex start: fileId={}, sheetName={}, rowIndex={}, allowRetry={}", safeString(fileID), safeString(sheetName), rowIndex, allowRetry);
            return sheetOperationsFacade.getAddRowWithIndexExtensionResponse(fileID, sheetName, rows, rowsFromCsv, rowIndex, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_4852fe2a-7ad3-49cf-b55c-a303ae8d21d7",
            name = "Add Row With Index Async",
            description = "Add Row With Index Async",
            area = "Sheet",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addRowWithIndexAsync(
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String sheetName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.Desc(name = "Rows From Csv", type = "[ File ]", required = false) List<File> rowsFromCsv,
            @Field(name = "Row Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S")}) Double rowIndex) {
        return audited("Add Row With Index Async", () -> {
            if (!Worksheets.isRowDataAvailable(rows, rowsFromCsv)) {
                throw new IllegalArgumentException(Constants.ROW_DATA_NOT_PROVIDED);
            }
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);

                    List<String> rowsStringOrCsvRows = stringUtil.getStringList(rows, rowsFromCsv);

                    restClientProvider.getRestClient(null).post(fileId + Constants.WORKBOOK_WORKSHEETS +
                                    sheetName + Constants.RANGE_ADDRESS +
                                    stringUtil.calculateAddressForInsert(rowIndex, rows,
                                            rowsFromCsv) + "')/insert",
                            Map.of(Constants.SHIFT, Constants.DOWN));
                    sheetFormattingService.clearRowFormatting(fileId, sheetName, rowIndex, rowsStringOrCsvRows.size());
                    StringUtil.appendCommas(rowsStringOrCsvRows);

                    List<FreeForm> result = worksheetService.addRows(null, fileId, sheetName, rowsStringOrCsvRows,
                            StringUtil.calculateAddressForInsert(rowIndex, rowsStringOrCsvRows));

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Result", "[ Composite ]", result);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (IOException cause) {
                    LOGGER.error("Error in addRowWithIndexAsync with fileId: {}, sheetName: {}, rowIndex: {}: {}",
                            fileId, sheetName, rowIndex, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_f1fdae78-a820-47a6-b3d2-00b44ec56f33",
            name = "Update Row With Index",
            description = "Update Row With Index",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    public ExtensionResponse updateRowWithIndex(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Row", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String row,
            @Field(name = "Row Index", type = "Number", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double rowIndex,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Update Row With Index", () -> {
            LOGGER.info("SheetArea.updateRowWithIndex start: fileId={}, sheetName={}, rowIndex={}, allowRetry={}", safeString(fileID), safeString(sheetName), rowIndex, allowRetry);
            return sheetOperationsFacade.getUpdateRowWithIndexExtensionResponse(fileID, sheetName, row, rowIndex, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_e7c64d77-3a3e-4be3-bfd9-e5de349d5365",
            name = "Get complete row info",
            description = "Retrieves all cell data and information for a specific row in the worksheet",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Row Data", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse getCompleteRowInfo(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field(name = "Row Index", type = "Number", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double rowIndex,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get complete row info", () -> {
            LOGGER.info("SheetArea.getCompleteRowInfo start: fileId={}, sheetName={}, rowIndex={}, allowRetry={}", safeString(fileId), safeString(sheetName), rowIndex, allowRetry);
            return sheetOperationsFacade.getGetCompleteRowInfoExtensionResponse(fileId, sheetName, rowIndex, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_b7f4544c-158e-4511-a07c-3720c466f124",
            name = "Delete Column With Index",
            description = "Deletes a specific column from the Excel worksheet by column name",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Boolean(name = "Deleted", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse deleteColumnWithIndex(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Column Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String columnName,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Delete Column With Index", () -> {
            LOGGER.info("SheetArea.deleteColumnWithIndex start: fileId={}, sheetName={}, columnName={}, allowRetry={}", safeString(fileID), safeString(sheetName), safeString(columnName), allowRetry);
            return sheetOperationsFacade.getDeleteColumnWithIndexExtensionResponse(fileID, sheetName, columnName, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_8b1ec8c9-6614-4edd-9a63-35612c9712d3",
            name = "Delete Row With Index",
            description = "Deletes a specific row from the Excel worksheet at the given row index",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Boolean(name = "Deleted", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse deleteRowWithIndex(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field(name = "Row Index", type = "Number", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double rowIndex,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Delete Row With Index", () -> {
            LOGGER.info("SheetArea.deleteRowWithIndex start: fileId={}, sheetName={}, rowIndex={}, allowRetry={}", safeString(fileID), safeString(sheetName), rowIndex, allowRetry);
            return sheetOperationsFacade.getDeleteRowWithIndexExtensionResponse(fileID, sheetName, rowIndex, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_10ae2fec-1397-48e4-bbaf-30edfa94f3d9",
            name = "Add Multiple Cell Data",
            description = "Add Multiple Cell Data",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Cell Data", type = "[ Composite ]", required = false)
    public ExtensionResponse addMultipleCellData(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Desc(name = "Cell ID Data", type = "[ { Cell Id: Text, Data: Text } ]") List<Map<String, Object>> cellIDData,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Add Multiple Cell Data", () -> {
            LOGGER.info("SheetArea.addMultipleCellData start: fileId={}, sheetName={}, allowRetry={}", safeString(fileID), safeString(sheetName), allowRetry);
            return sheetOperationsFacade.getAddMultipleCellDataExtensionResponse(fileID, sheetName, cellIDData, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_006a6d7e-5622-4a9b-8805-461c81293761",
            name = "Add Multiple Cell Data Async",
            description = "Add Multiple Cell Data Async",
            area = "Sheet",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addMultipleCellDataAsync(
            @Field.Text(name = "File ID", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileID,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String sheetName,
            @Field.Desc(name = "Cell ID Data", type = "[ { Cell Id: Text, Data: Text } ]") List<Map<String, Object>> cellIDData) {
        return audited("Add Multiple Cell Data Async", () -> {
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);

                    List<FreeForm> allCellData = new ArrayList<>();

                    for (Map eachMap : cellIDData) {
                        String cellId = (String) eachMap.get("Cell Id");
                        if (!Worksheets.isValidCellId(cellId)) {
                            throw new IllegalArgumentException(Constants.INVALID_CELL_ID_IS_GIVEN + cellId);
                        }
                        String dataToBeAdded = (String) eachMap.get(Constants.DATA);
                        allCellData.add(this.worksheetService.addDataToCell(null, fileID, sheetName, cellId, dataToBeAdded));
                    }

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Result", "[ Composite ]", allCellData);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (Exception e) {
                    LOGGER.error("Error in addMultipleCellDataAsync with fileID: {}, sheetName: {}: {}",
                            fileID, sheetName, e.getMessage(), e);
                    throw new RuntimeException(e);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_b1611ba3-2b1b-42e8-9b77-e4ffe45d9506",
            name = "Find Value",
            description = "Find Value",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Cell IDs", type = "[ Text ]", required = false)
    public ExtensionResponse findValue(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Lookup Value", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String lookupValue,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Find Value", () -> {
            LOGGER.info("SheetArea.findValue start: fileId={}, sheetName={}, lookupValue={}, allowRetry={}", safeString(fileID), safeString(sheetName), safeString(lookupValue), allowRetry);
            return sheetOperationsFacade.getFindValueExtensionResponse(fileID, sheetName, lookupValue, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_86dcd7ba-3bd3-403c-9f58-ef678e31543f",
            name = "Get Excel File Information",
            description = "Get Excel File Information",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Excel File", type = "Entity(Excel File)", required = false)
    public ExtensionResponse getExcelFileInformation(
            @Field.Text(name = "Path", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String path,
            @Field.Text(name = "Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String name,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Get Excel File Information", () -> {
            LOGGER.info("SheetArea.getExcelFileInformation start: path={}, name={}, allowRetry={}", safeString(path), safeString(name), allowRetry);
            return sheetOperationsFacade.getGetExcelFileInformationExtensionResponse(path, name, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_80a0e5d8-2267-431d-aec8-6441f7dd4dbc",
            name = "Get Sheets Information",
            description = "Return names and ids of the sheets defined in excel.",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Sheets", type = "[ Entity(Excel Sheet) ]", required = false)
    public ExtensionResponse getSheetsInformation(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Sheets Information", () -> {
            LOGGER.info("SheetArea.getSheetsInformation start: fileId={}, allowRetry={}", safeString(fileId), allowRetry);
            return sheetOperationsFacade.getGetSheetsInformationExtensionResponse(fileId, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_810479c1-5fba-4337-a54a-18888d1c04ec",
            name = "Add Data To Cell",
            description = "Add Data To Cell",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Cell Information", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse addDataToCell(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Cell ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String cellID,
            @Field.Text(name = "Data", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String data,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Add Data To Cell", () -> {
            LOGGER.info("SheetArea.addDataToCell start: fileId={}, sheetName={}, cellId={}, allowRetry={}", safeString(fileID), safeString(sheetName), safeString(cellID), allowRetry);
            return sheetOperationsFacade.getAddDataToCellExtensionResponse(fileID, sheetName, cellID, data, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_5c0e41ef-759e-45d2-b977-14853d75f979",
            name = "Execute Formula",
            description = "Provided Formula will be executed and will return result.",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Formula Result", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse executeFormula(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Formula", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String formula,
            @Field.Text(name = "Formula Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String formulaName,
            @Field.Text(name = "Comment", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String comment,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Execute Formula", () -> {
            LOGGER.info("SheetArea.executeFormula start: fileId={}, formula={}, formulaName={}, allowRetry={}", safeString(fileId), safeString(formula), safeString(formulaName), allowRetry);
            return sheetOperationsFacade.getExecuteFormulaExtensionResponse(fileId, formula, formulaName, comment, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_1d637b17-dee5-4095-a627-d774e0621d1b",
            name = "Create Sheet",
            description = "Create new Sheet in Excel File",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Sheet", type = "Entity(Excel Sheet)", required = false)
    public ExtensionResponse createSheet(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "New Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String newSheetName,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Create Sheet", () -> {
            LOGGER.info("SheetArea.createSheet start: fileId={}, newSheetName={}, allowRetry={}", safeString(fileId), safeString(newSheetName), allowRetry);
            return sheetOperationsFacade.getCreateSheetExtensionResponse(fileId, newSheetName, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_29c0c32f-7d07-47d4-9e04-5198f6c2297b",
            name = "Clear Sheet For Range",
            description = "Clear Sheet For Range",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Boolean(name = "Cleared", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse clearSheetForRange(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Cell Range", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String cellRange,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Clear Sheet For Range", () -> {
            LOGGER.info("SheetArea.clearSheetForRange start: fileID={}, sheetName={}, cellRange={}, allowRetry={}",
                    safeString(fileID), safeString(sheetName), safeString(cellRange), allowRetry);
            return sheetOperationsFacade.getClearSheetForRangeExtensionResponse(fileID, sheetName, cellRange, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_d2d313fb-51de-4d80-8ad3-99c1889cf433",
            name = "Vlookup",
            description = "Vlookup",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Lookup Result", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse vlookup(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Lookup Value", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String lookupValue,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Range", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String range,
            @Field(name = "Column Index", type = "Number", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double columnIndex,
            @Field.Boolean(name = "Exact Match", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean exactMatch,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Vlookup", () -> {
            LOGGER.info("SheetArea.vlookup start: fileID={}, lookupValue={}, sheetName={}, range={}, columnIndex={}, exactMatch={}, allowRetry={}",
                    safeString(fileID), safeString(lookupValue), safeString(sheetName), safeString(range), columnIndex, exactMatch, allowRetry);
            return sheetOperationsFacade.getVlookupExtensionResponse(fileID, lookupValue, sheetName, range, columnIndex, exactMatch, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_36a7b041-742b-429e-85d2-0fdfde4dea7d",
            name = "Get Rows For Range",
            description = "Retrieves all rows of data within the specified cell range from an Excel worksheet",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse getRowsForRange(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Range", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String range,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Rows For Range", () -> {
            LOGGER.info("SheetArea.getRowsForRange start: fileID={}, sheetName={}, range={}, allowRetry={}", safeString(fileID), safeString(sheetName), safeString(range), allowRetry);
            return sheetOperationsFacade.getGetRowsForRangeExtensionResponse(fileID, sheetName, range, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_a1b2c3d4-e5f6-7890-abcd-ef1234567891",
            name = "Get Rows For Range By Path",
            description = "Retrieves all rows of data within the specified cell range from an Excel worksheet, locating the file by its SharePoint folder path instead of a File ID. Source Path should be the full path to the Excel file, e.g. /Engineering/Shared Documents/Reports/Q1_Report.xlsx. URL-encoded paths (e.g. Shared%20Documents) are supported.",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse getRowsForRangeByPath(
            @Field.Text(name = "Source Path", required = true,
                    description = "Full SharePoint path to the Excel file, e.g. /Engineering/Shared Documents/Reports/Q1_Report.xlsx. URL-encoded paths supported.",
                    attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sourcePath,
            @Field.Text(name = "Sheet Name", required = true,
                    attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Range", required = true,
                    attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String range,
            @Field.PickOne(name = "File Search", required = true,
                    values = {"My Drive", "Site Files"},
                    attributes = {@Attribute(name = "visualWidth", value = "S"),
                            @Attribute(name = "toolTip", value = "'My Drive: file is in your own personal OneDrive (provide full folder path). Site Files: file lives inside a SharePoint site you are a member of (path must start with site name and library).'")},
                    options = {}) String fileSearch,
            @Field.Boolean(name = "Allow Retry", required = false,
                    attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Rows For Range By Path", () -> {
            LOGGER.info("SheetArea.getRowsForRangeByPath start: sourcePath={}, sheetName={}, range={}, fileSearch={}",
                    safeString(sourcePath), safeString(sheetName), safeString(range), safeString(fileSearch));
            ExcelFile file = driveService.getExcelFileByPath(sourcePath, fileSearch);
            LOGGER.info("SheetArea.getRowsForRangeByPath resolved: fileId={}, driveId={}", file.fileId, file.driveId);
            return sheetOperationsFacade.getGetRowsForRangeExtensionResponse(file.driveId, file.fileId, sheetName, range, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_ff9c58ad-39a2-4a6f-aced-b713e9f69c6c",
            name = "Add Formula To Cell",
            description = "Adds an Excel formula to a specific cell in the worksheet",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Cell Information", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse addFormulaToCell(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Cell ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String cellID,
            @Field.Text(name = "Formula", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String formula,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Add Formula To Cell", () -> {
            LOGGER.info("SheetArea.addFormulaToCell start: fileID={}, sheetName={}, cellID={}, formula={}, allowRetry={}",
                    safeString(fileID), safeString(sheetName), safeString(cellID), safeString(formula), allowRetry);
            return sheetOperationsFacade.getAddFormulaToCellExtensionResponse(fileID, sheetName, cellID, formula, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_926ddadc-8a15-4811-8516-b31ed923fd96",
            name = "Get Min Max Value From Row Data By Column Range",
            description = "Retrieves the minimum and maximum numeric values from a specified range of columns and their corresponding key column values",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Result", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse getMinMaxValueFromRowDataByColumnRange(
            @Field(name = "File Id", type = "Text", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field(name = "Sheet Name", type = "Text", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field(name = "Key Column Name", type = "Text", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String keyColumnName,
            @Field(name = "Column Range", type = "Text", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String columnRange,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Min Max Value From Row Data By Column Range", () -> {
            LOGGER.info("SheetArea.getMinMaxValueFromRowDataByColumnRange start: fileId={}, sheetName={}, keyColumnName={}, columnRange={}, allowRetry={}",
                    safeString(fileId), safeString(sheetName), safeString(keyColumnName), safeString(columnRange), allowRetry);
            return sheetOperationsFacade.getGetMinMaxValueFromRowDataByColumnRangeExtensionResponse(fileId, sheetName, keyColumnName, columnRange, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_6aa5b93f-6a9d-4054-ab89-3855da644c3f",
            name = "Summation",
            description = "Calculates the sum of numeric values within a specified range in an Excel worksheet",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Sum", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse summation(
            @Field.Text(name = "File ID", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileID,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Range", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String range,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Summation", () -> {
            LOGGER.info("SheetArea.summation start: fileID={}, sheetName={}, range={}, allowRetry={}", safeString(fileID), safeString(sheetName), safeString(range), allowRetry);
            return sheetOperationsFacade.getSummationExtensionResponse(fileID, sheetName, range, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_3a8d0518-e94a-4380-96df-6cbd26205261",
            name = "Upload And Convert To PDF",
            description = "Uploads an Excel file to OneDrive and converts it to PDF format",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.File(name = "PDF File", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    @Field(name = "File Id", type = "Text", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse uploadFileAndConvertToPDF(
            @Field.File(name = "Excel File", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) File excelFile,
            @Field(name = "File Path", type = "Text", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String filePath,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Upload And Convert To PDF", () -> {
            LOGGER.info("SheetArea.uploadFileAndConvertToPDF start: excelFile={}, filePath={}, allowRetry={}", excelFile != null ? excelFile : "null", safeString(filePath), allowRetry);
            return sheetOperationsFacade.getUploadFileAndConvertToPDFExtensionResponse(excelFile, filePath, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_8adce3d3-edb1-4dae-b9b1-9147399abb8b",
            name = "Convert To PDF",
            description = "Converts an Excel file from OneDrive to PDF format while maintaining the same layout",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.File(name = "PDF File", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse convertToPDF(
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Convert To PDF", () -> {
            LOGGER.info("SheetArea.convertToPDF start: fileId={}, allowRetry={}", safeString(fileId), allowRetry);
            return sheetOperationsFacade.getConvertToPDFExtensionResponse(fileId, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_f74389ac-3464-4fe8-a4c9-d7244f73d07c",
            name = "Get Column Values For Keys In Sheet",
            description = "Retrieves column values based on key rows and columns from a specified range in the worksheet",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Data", type = "[ Composite ]", required = false)
    public ExtensionResponse getColumnValuesForKeysInSheet(
            @Field.Text(name = "Sheet  Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Range", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String range,
            @Field.Desc(name = "Key Rows", type = "[ Text ]", required = true) List<String> keyRows,
            @Field.Desc(name = "Key Columns", type = "[ Text ]", required = true) List<String> keyColumns,
            @Field.Boolean(name = "Get All Records", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean getAllRecords,
            @Field(name = "Column Number", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Double columnNumber,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Get Column Values For Keys In Sheet", () -> {
            LOGGER.info("SheetArea.getColumnValuesForKeysInSheet start: fileId={}, sheetName={}, range={}, keyRows={}, keyColumns={}, getAllRecords={}, columnNumber={}, allowRetry={}",
                    safeString(fileId), safeString(sheetName), safeString(range), keyRows, keyColumns, getAllRecords, columnNumber, allowRetry);
            return sheetOperationsFacade.getGetColumnValuesForKeysInSheetExtensionResponse(fileId, sheetName, range, keyRows, keyColumns, getAllRecords, columnNumber, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_c33dd315-7eb1-41de-a42d-5cc7b274539e",
            name = "Add Rows With The List Of Entities",
            description = "Add Rows With The List Of Entities",
            area = "Sheet",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Final Data", type = "[ Composite ]", required = false)
    public List<FreeForm> addRowsWithTheListOfEntities(
            @Field.Text(name = "File ID", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileID,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String sheetName,
            @Field.Desc(name = "Data", type = "[ Entity ]") List<EntityValue> data,
            @Field.Desc(name = "Exclude Columns", type = "[ Text ]", required = false) List<String> excludeColumns) {
        return audited("Add Rows With The List Of Entities", () -> {
            if (StringUtil.isDataNullOrEmpty(data)) {
                return List.of((FreeForm) Constants.FREE_FORM.put("Error", "Text", "No data provided. Please provide valid entity data to add to the worksheet."));
            }
            EntityDefinition entityDefinition = registry.getEntityDefinition(data.getFirst());
            // Retrieve the used range of the sheet to determine if it's empty or not.
            String usedRangeForSheet = worksheetOperation.getUsedRangeForSheet(null,
                    fileID, sheetName);

            return this.worksheetOperation.addRowsToSheetWithEntity(fileID, sheetName, data, usedRangeForSheet, null, excludeColumns, entityDefinition);
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
