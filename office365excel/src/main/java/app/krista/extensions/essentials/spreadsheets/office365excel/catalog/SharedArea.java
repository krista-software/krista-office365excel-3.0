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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Drive;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.SharedOperationsFacade;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.util.EventHandler;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

@Domain(id = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24",
        name = "Spreadsheets",
        ecosystemId = "catEntryEcosystem_84b53163-327b-4b1b-8c96-9334d292f9f5",
        ecosystemName = "Essentials",
        ecosystemVersion = "60a6d210-d07b-4b1a-b07f-2abb0c42648c")
public class SharedArea {

    private static final Logger LOGGER = LoggerFactory.getLogger(SharedArea.class);

    private final DriveService driveService;
    private final TableService tableService;
    private final RestClientProvider restClientProvider;
    private final WorkbookService workbookService;
    private final WorksheetService worksheetService;
    private final EventHandler eventHandler;
    private final RequestContext requestContext;
    ExecutorService executorService = Executors.newSingleThreadExecutor();
    private final AuthorizationContext authorizationContext;
    private final Entities registry;
    private final WorksheetOperation worksheetOperation;
    private final StringUtil stringUtil;
    private final SheetFormattingService sheetFormattingService;
    private final SharedOperationsFacade sharedOperationsFacade;
    private final Office365ExcelAuditService auditService;
    private final String invokerId;
    private final GraphServiceClientProviderFactory providerFactory;

    @Inject
    public SharedArea(DriveService driveService, RestClientProvider restClientProvider, RequestContext requestContext,
                      EventHandler eventHandler, AuthorizationContext authorizationContext, Entities entities,
                      WorkbookService workbookService, TableService tableService, WorksheetService worksheetService,
                      WorksheetOperation worksheetOperation, StringUtil stringUtil, SheetFormattingService sheetFormattingService,
                      SharedOperationsFacade sharedOperationsFacade, Office365ExcelAuditService auditService,
                      GraphServiceClientProviderFactory providerFactory, Invoker invoker) {
        this.driveService = driveService;
        this.restClientProvider = restClientProvider;
        this.eventHandler = eventHandler;
        this.registry = entities;
        this.workbookService = workbookService;
        this.worksheetService = worksheetService;
        this.tableService = tableService;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.worksheetOperation = worksheetOperation;
        this.stringUtil = stringUtil;
        this.sheetFormattingService = sheetFormattingService;
        this.sharedOperationsFacade = sharedOperationsFacade;
        this.auditService = auditService;
        this.providerFactory = providerFactory;
        this.invokerId = invoker.getInvokerId();
    }

    @CatalogRequest(
            id = "localDomainRequest_e045b09c-2ae6-4ed8-8d54-7275bcf85c84",
            name = "Add Rows to Table",
            description = "Accepts drive ID, file ID, table name, rows, and rows from csv file and returns rows. Rows, and rows from csv file are optional inputs.",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse addRowsToTable(
            @Field.Text(name = "Drive Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.File(name = "Rows from Csv File", multipleFileUpload = true, required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) List<File> rowsFromCsvFile,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {
        return audited("Add Rows to Table", () -> {
            LOGGER.info("CatalogRequest addRowsToTable start with driveId: {}, fileId: {}, tableName: {}, allowRetry={}",
                    safeString(driveId), safeString(fileId), safeString(tableName), allowRetry);
            return sharedOperationsFacade.getExtensionResponse(driveId, fileId, tableName, rows, rowsFromCsvFile, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_a463aac5-a65e-469e-beb3-88aad159188b",
            name = "Add Rows to Table Async",
            description = "Add Rows to Table Async",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addRowsToTableAsync(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileId,
            @Field.Text(name = "Table Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String tableName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.Desc(name = "Rows From Csv", type = "[ File ]", required = false) List<File> rowsFromCsv) {
        if (!Worksheets.isRowDataAvailable(rows, rowsFromCsv)) {
            throw new IllegalArgumentException("Row data not provided");
        }
        Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);

                    List<FreeForm> result = workbookService.addRowsInTable(driveId, fileId, tableName, rows, rowsFromCsv, null, Constants.COMMA);

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Sheet Data", "[ Composite ]", result);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (Exception cause) {
                    LOGGER.error("Error in addRowsToTableAsync with driveId: {}, fileId: {}, tableName: {}: {}",
                            driveId, fileId, tableName, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

        return taskId;
    }

    @CatalogRequest(
            id = "localDomainRequest_a82488b8-3dc2-47f2-b7d8-1a220027ad52",
            name = "Add Rows to Sheet",
            description = "Accepts drive ID, file ID, sheet name, rows, and rows from file and returns sheet data. Rows, and rows from file are optional inputs.",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    public ExtensionResponse addRowsToSheet(
            @Field.Text(name = "Drive Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field.File(name = "Rows from File", multipleFileUpload = true, required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) List<File> rowsFromFile,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Add Rows to Sheet", () -> {
            LOGGER.info("CatalogRequest addRowsToSheet start with driveId: {}, fileId: {}, sheetName: {}, allowRetry={}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), allowRetry);
            return sharedOperationsFacade.getSheetExtensionResponse(driveId, fileId, sheetName, rows, rowsFromFile, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_5179ebac-9ac2-4d1c-b19b-9090ecf2fed0",
            name = "Range Sum",
            description = "Accepts drive ID, file ID, sheet name, and range and returns sum of given range",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Sum", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse rangeSum(
            @Field.Text(name = "Drive Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Sheet Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String sheetName,
            @Field.Text(name = "Range", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String range,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Range Sum", () -> {
            LOGGER.info("CatalogRequest rangeSum start with driveId: {}, fileId: {}, sheetName: {}, range: {}, allowRetry={}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(range), allowRetry);
            return sharedOperationsFacade.getRangeSumExtensionResponse(driveId, fileId, sheetName, range, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_c4e8b959-aa4c-4b74-89b8-f98ebbcf7c15",
            name = "Sheet Vlookup",
            description = "Accepts drive ID, file ID, lookup value, sheet name, range, column index, and exact match and returns vlookup result",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field(name = "Vlookup Result", type = "FreeForm", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {})
    public ExtensionResponse sheetVlookup(
            @Field.Text(name = "Drive Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
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
        return audited("Sheet Vlookup", () -> {
            LOGGER.info("CatalogRequest sheetVlookup start with driveId: {}, fileId: {}, lookupValue: {}, sheetName: {}, range: {}, columnIndex: {}, exactMatch: {}, allowRetry={}",
                    safeString(driveId), safeString(fileId), safeString(lookupValue), safeString(sheetName), safeString(range), columnIndex, exactMatch, allowRetry);
            return sharedOperationsFacade.getSheetVlookupExtensionResponse(driveId, fileId, lookupValue, sheetName, range, columnIndex, exactMatch, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_f17b7b62-3218-4837-bc59-49cf9c49349d",
            name = "Get Table Rows",
            description = "Accepts drive ID, file ID, and table name as input and returns table rows",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Table Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse getTableRows(
            @Field.Text(name = "Drive Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Table Rows", () -> {
            LOGGER.info("CatalogRequest getTableRows start with driveId: {}, fileId: {}, tableName: {}, allowRetry={}",
                    safeString(driveId), safeString(fileId), safeString(tableName), allowRetry);
            return sharedOperationsFacade.getTableRowsExtensionResponse(driveId, fileId, tableName, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_9a51766e-f0b5-4e2f-8a09-1fc1b5785364",
            name = "Get Shared Excel Files",
            description = "Returns list of file.",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Files", type = "[ Entity(Excel File) ]")
    public List<ExcelFile> getSharedExcelFiles(
            @Field.Text(name = "Drive Id", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId) {
        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Shared Excel Files", () -> {
            if (driveId != null && !driveId.isBlank()) {
                LOGGER.info("getSharedExcelFiles: using direct drive enumeration for driveId={}", driveId);
                return this.driveService.getExcelFilesFromDrive(driveId);
            }
            return this.driveService.getSharedExcelFilesInformation(null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_7f3a8b2e-5d19-4c6a-b8e1-9a4d2f7c3e05",
            name = "Search All Shared Excel Files in Drive",
            description = "Searches for all Excel files (.xlsx, .xls, .csv, .tsv) recursively across all folders in a drive using the Microsoft Graph search API.",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Files", type = "[ Entity(Excel File) ]")
    public List<ExcelFile> searchExcelFilesInDrive(
            @Field.Text(name = "Drive Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.PickOne(name = "File Format", values = {".xlsx", ".xls", ".csv", ".tsv"}, required = true, description = "", attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileFormat) {
        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Search Excel Files in Drive", () -> {
            LOGGER.info("CatalogRequest searchExcelFilesInDrive start with driveId: {}, fileFormat: {}", safeString(driveId), safeString(fileFormat));
            return this.driveService.searchExcelFilesInDrive(driveId, fileFormat);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_b3d27624-7d1c-4f9b-903d-fac69e5ae82a",
            name = "Add Data to provided Cell",
            description = "Add Data to provided Cell",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Cell Information", type = "FreeForm", required = false, attributes = {@Attribute(name = "toolTip", value = "'Information about the cell after data was added'")})
    public ExtensionResponse addDataToProvidedCell(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the cell'")}) String sheetName,
            @Field.Text(name = "Cell Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the cell where data will be added (e.g., A1, B2)'")}) String cellId,
            @Field.Text(name = "Data", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The data to be added to the specified cell'")}) String data) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Add Data to provided Cell", () -> {
            LOGGER.info("CatalogRequest addDataToProvidedCell start with driveId: {}, fileId: {}, sheetName: {}, cellId: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(cellId));
            return sharedOperationsFacade.getAddDataToCellExtensionResponse(driveId, fileId, sheetName, cellId, data, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_5a3cda6b-aec1-4acf-a715-ab5e383d237a",
            name = "Delete Column For Shared File",
            description = "Delete a particular column from the given table from the spreadsheet.",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Is Deleted", type = "Switch", required = false, attributes = {@Attribute(name = "toolTip", value = "'Indicates whether the column was successfully deleted'")})
    public ExtensionResponse deleteColumn(
            @Field.Text(name = "Drive ID", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveID,
            @Field.Text(name = "File ID", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the table'")}) String fileID,
            @Field.Text(name = "Table Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the table containing the column to delete'")}) String tableName,
            @Field.Text(name = "Column Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the column to be deleted'")}) String columnName) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Delete Column For Shared File", () -> {
            LOGGER.info("CatalogRequest deleteColumn start with driveId: {}, fileId: {}, tableName: {}, columnName: {}",
                    safeString(driveID), safeString(fileID), safeString(tableName), safeString(columnName));
            return sharedOperationsFacade.getDeleteColumnExtensionResponse(driveID, fileID, tableName, columnName, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_c2fd7c21-84cf-45e2-a167-8610ae8bba68",
            name = "Get Column Values For Keys",
            description = "Get column values for key rows and lookup columns.",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Result", type = "[ Composite ]", required = false)
    public ExtensionResponse getColumnValuesForKeys(
            @Field.Text(name = "Drive Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId,
            @Field.Text(name = "File Id", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String fileId,
            @Field.Text(name = "Table Name", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String tableName,
            @Field.Desc(name = "Key Rows", type = "[ Text ]", required = true) List<String> keyRows,
            @Field.Desc(name = "Lookup Column", type = "[ Text ]", required = true) List<String> lookupColumn,
            @Field.Boolean(name = "Get All Record", required = true, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean getAllRecord,
            @Field.Boolean(name = "Allow Retry", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Column Values For Keys", () -> {
            LOGGER.info("CatalogRequest getColumnValuesForKeys start with driveId: {}, fileId: {}, tableName: {}, keyRows: {}, lookupColumn: {}, getAllRecord: {}, allowRetry={}",
                    safeString(driveId), safeString(fileId), safeString(tableName), keyRows, lookupColumn, getAllRecord, allowRetry);
            return sharedOperationsFacade.getColumnValuesForKeysExtensionResponse(driveId, fileId, tableName, keyRows, lookupColumn, getAllRecord, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_b39195e1-e54a-47f2-bb35-beae5885d382",
            name = "Get Rows for Range",
            description = "Get Rows for Range",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    public ExtensionResponse getRowsForRange(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the range'")}) String sheetName,
            @Field.Text(name = "Range", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The Excel range to retrieve rows from (e.g., A1:D10)'")}) String range,
            @Field.Boolean(name = Constants.ENABLE_ERROR_HANDLING, required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) Boolean allowRetry) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Rows for Range", () -> {
            LOGGER.info("CatalogRequest getRowsForRange start with driveId: {}, fileId: {}, sheetName: {}, range: {}, allowRetry={}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(range), allowRetry);
            return sharedOperationsFacade.getRowsForRangeExtensionResponse(driveId, fileId, sheetName, range, allowRetry);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_d7a2f3c1-8b5e-4e91-a6d4-3f9c0e7b2a18",
            name = "Get Rows for Range Async",
            description = "Get Rows for Range Async",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String getRowsForRangeAsync(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the range'")}) String sheetName,
            @Field.Text(name = "Range", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The Excel range to retrieve rows from (e.g., A1:D10)'")}) String range) {
        return audited("Get Rows for Range Async", () -> {
            LOGGER.info("CatalogRequest getRowsForRangeAsync start with driveId: {}, fileId: {}, sheetName: {}, range: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(range));
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            final String taskId = UUID.randomUUID().toString();
            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);

                    List<FreeForm> result = worksheetService.getRows(driveId, fileId, sheetName, range);

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Rows", "[ Composite ]", result);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (Exception cause) {
                    LOGGER.error("Error in getRowsForRangeAsync with driveId: {}, fileId: {}, sheetName: {}, range: {}: {}",
                            driveId, fileId, sheetName, range, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_d7a2f3c1-8b5e-4e91-a6d4-3f9c0e7b2a18_result",
            name = "Get Rows for Range Result",
            description = "Get the result of the asynchronous Get Rows for Range operation.",
            area = "Shared",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field.Desc(name = "Rows", type = "[ Composite ]", required = false)
    public List<FreeForm> getRowsForRangeResult(
            @Field(name = "eventName", type = "Text") String eventName,
            @Field(name = "eventData", type = "FreeForm") FreeForm eventData,
            @Field(name = "taskID", type = "Text") String taskID) {

        return audited("Get Rows for Range Result", () -> {
            if (eventName.equals(taskID)) {
                return (List<FreeForm>) eventData.get("Rows");
            }
            throw new IllegalStateException("Invalid task ID");
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_c96220d2-fddd-48e2-8c53-fc78d7518779",
            name = "Execute Formula For Shared File",
            description = "Formula will be executed and returns a result.",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Formula Result", type = "FreeForm", required = false, attributes = {@Attribute(name = "toolTip", value = "'The result of executing the formula'")})
    public ExtensionResponse executeFormulaForSharedFile(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file where the formula will be executed'")}) String fileId,
            @Field.Text(name = "Formula", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The Excel formula to execute (e.g., =SUM(A1:A10))'")}) String formula,
            @Field.Text(name = "Formula Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'A name to identify this formula'")}) String formulaName,
            @Field.Text(name = "Comment", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'Optional comment about the formula'")}) String comment) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Execute Formula For Shared File", () -> {
            LOGGER.info("CatalogRequest executeFormulaForSharedFile start with driveId: {}, fileId: {}, formula: {}, formulaName: {}",
                    safeString(driveId), safeString(fileId), safeString(formula), safeString(formulaName));
            return sharedOperationsFacade.getExecuteFormulaExtensionResponse(driveId, fileId, formula, formulaName, comment, null);
        });
    }

    @CatalogRequest(id = "localDomainRequest_164c1033-d2cb-4a30-beb2-c079f16a9e59",
            name = "Vlookup in Table",
            description = "Vlookup in Table",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field(name = "Vlookup Result", type = "FreeForm", required = false, attributes = {@Attribute(name = "toolTip", value = "'The result of the VLOOKUP operation'")})
    public ExtensionResponse vlookupInTable(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the table'")}) String fileId,
            @Field.Text(name = "Lookup Value", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The value to look up in the first column of the table'")}) String lookupValue,
            @Field.Text(name = "Table Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the table to perform lookup in'")}) String tableName,
            @Field(name = "Column Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The index of the column to return the value from'")}) Double columnIndex,
            @Field.Boolean(name = "Exact Match", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'If true, requires an exact match; if false, finds the closest match'")}) Boolean exactMatch) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Vlookup in Table", () -> {
            LOGGER.info("CatalogRequest vlookupInTable start with driveId: {}, fileId: {}, lookupValue: {}, tableName: {}, columnIndex: {}, exactMatch: {}",
                    safeString(driveId), safeString(fileId), safeString(lookupValue), safeString(tableName), columnIndex, exactMatch);
            return sharedOperationsFacade.getVlookupInTableExtensionResponse(driveId, fileId, lookupValue, tableName, columnIndex, exactMatch, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_a49685dd-6d97-4aef-b372-12c1e9b7efb8",
            name = "Table Information For The Shared File Async",
            description = "Get the information about the table with the sheet name",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field(name = "taskID", type = "Text", required = false)
    public String tableInformationForTheSharedFileAsync(
            @Field(name = "fileId", type = "Text") String fileId,
            @Field(name = "driveId", type = "Text") String driveId) {
        return audited("Table Information For The Shared File Async", () -> {
            final String taskId = UUID.randomUUID().toString();
            boolean invokeAsUser = requestContext.invokeAsUser();
            String accountID = invokeAsUser ? authorizationContext.getAuthorizedAccount().getAccountId()
                    : restClientProvider.getOffice365ExcelAttributes().getEmailId();
            Map<Class<?>, Object> threadLocals = ThreadLocalProxy.getAll();

            executorService.submit(() -> {
                try {
                    ThreadLocalProxy.setAll(threadLocals);
                    List<Table> tableNames = this.tableService.getTableNames(accountID, invokeAsUser, fileId, driveId);
                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Data", "[ Entity(Table) ]", tableNames);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (Exception cause) {
                    LOGGER.error("Error in tableInformationForTheSharedFileAsync with driveId: {}, fileId: {}: {}",
                            driveId, fileId, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_cba9817f-b70b-47c9-b878-6aed9a0e09f2",
            name = "Get Tables Information For Shared File",
            description = "Accept task ID as input and return mails. Get this tasks ID from tableInfoForTheSharedFileAsync request",
            area = "Shared",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field.Desc(name = "Tables Details", type = "[ Entity(Table) ]", required = false)
    public List<Table> getTablesInformationForSharedFile(
            @Field(name = "eventName", type = "Text") String eventName,
            @Field(name = "eventData", type = "FreeForm") FreeForm eventData,
            @Field(name = "taskID", type = "Text") String taskID) {

        return audited("Get Tables Information For Shared File", () -> {
            if (eventName.equals(taskID)) {
                return (List<Table>) eventData.get("Data");
            }
            throw new IllegalStateException("Invalid task ID");
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_e7e64e89-9f39-45b4-881d-181682098f33",
            name = "Columns Sum",
            description = "Perform sum on a column or columns of the table.",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Sum", type = "Number", attributes = {@Attribute(name = "toolTip", value = "'The sum of values in the specified columns'")})
    public ExtensionResponse columnsSum(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the table'")}) String fileId,
            @Field.Text(name = "Table Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the table containing the columns to sum'")}) String tableName,
            @Field.Desc(name = "Columns", type = "[ Text ]") List<String> columns) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Columns Sum", () -> {
            LOGGER.info("CatalogRequest columnsSum start with driveId: {}, fileId: {}, tableName: {}, columns: {}",
                    safeString(driveId), safeString(fileId), safeString(tableName), columns);
            return sharedOperationsFacade.getColumnsSumExtensionResponse(driveId, fileId, tableName, columns, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_c96454cc-787d-49f2-9b57-e38449bc5e7a",
            name = "Search",
            description = "Search Data in Sheet and returns corresponding Cell Ids.",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Cell Ids", type = "[ Text ]", required = false)
    public ExtensionResponse search(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet to search in'")}) String sheetName,
            @Field.Text(name = "Data to Search", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The text or value to search for in the sheet'")}) String dataToSearch) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Search", () -> {
            LOGGER.info("CatalogRequest search start with driveId: {}, fileId: {}, sheetName: {}, dataToSearch: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(dataToSearch));
            return sharedOperationsFacade.getSearchExtensionResponse(driveId, fileId, sheetName, dataToSearch, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_c601297d-f88d-44b0-a6ca-b1b27ad2c81b",
            name = "Cell Data",
            description = "Get Cell Information along with value.",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field(name = "Cell Information", type = "FreeForm", required = false)
    public ExtensionResponse cellData(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the cell'")}) String sheetName,
            @Field.Text(name = "Cell Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the cell to retrieve data from'")}) String cellId) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Cell Data", () -> {
            LOGGER.info("CatalogRequest cellData start with driveId: {}, fileId: {}, sheetName: {}, cellId: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(cellId));
            return sharedOperationsFacade.getCellDataExtensionResponse(driveId, fileId, sheetName, cellId, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_e5b1befb-b2d1-432e-a1ac-d6be09f5309e",
            name = "Get All Shared Drives Information",
            description = "It Gives all Shared Drives Information.",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM,
            tool = true)
    @Field.Desc(name = "Drives", type = "[ Entity(Drive) ]")
    public List<Drive> getAllSharedDrivesInformation(
            @Field.Text(name = "Drive Id", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")}, options = {}) String driveId) {
        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get All Shared Drives Information", () -> {
            if (driveId != null && !driveId.isBlank()) {
                LOGGER.info("getAllSharedDrivesInformation: also fetching drive by id={}", driveId);
                List<Drive> drives = new ArrayList<>(this.driveService.getAllSharedDrivesInformation());
                List<Drive> byId = this.driveService.getDriveById(driveId);
                for (Drive d : byId) {
                    boolean alreadyPresent = drives.stream().anyMatch(existing -> d.driveId != null && d.driveId.equals(existing.driveId));
                    if (!alreadyPresent) {
                        drives.add(d);
                    }
                }
                return drives;
            }
            return this.driveService.getAllSharedDrivesInformation();
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_aa9b37c0-8d82-4cdc-9481-01cd93c4bbc8",
            name = "Get Shared Excel File",
            description = "Retrieve the shared file from specified shared folder name",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Shared Excel File", type = "Entity(Excel File)", required = false)
    public ExtensionResponse getSharedExcelFile(
            @Field.Text(name = "Shared Folder Name", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the folder containing the Excel file (optional)'")}) String sharedFolderName,
            @Field.Text(name = "Shared File Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the Excel file to retrieve'")}) String sharedFileName) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get Shared Excel File", () -> {
            LOGGER.info("CatalogRequest getSharedExcelFile start with sharedFolderName: {}, sharedFileName: {}",
                    safeString(sharedFolderName), safeString(sharedFileName));
            return sharedOperationsFacade.getSharedExcelFileExtensionResponse(sharedFolderName, sharedFileName, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_a438702d-ad8c-4f53-bf5b-405035dc8ef5",
            name = "Convert To PDF",
            description = "This method will take one drive excel file Id and drive Id as a user "
                    + "input and convert into PDF with same format",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.File(name = "PDF File", required = false)
    public ExtensionResponse convertToPDF(
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file to convert to PDF'")}) String fileId,
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Convert To PDF", () -> {
            LOGGER.info("CatalogRequest convertToPDF start with fileId: {}, driveId: {}", safeString(fileId), safeString(driveId));
            return sharedOperationsFacade.getConvertToPdfExtensionResponse(driveId, fileId, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_332de10e-1308-4ed9-a4c9-d92e5efc2781",
            name = "Delete Column With Index",
            description = "Delete Column With Index",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Boolean(name = "Deleted", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public ExtensionResponse deleteColumnWithIndex(
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the column to delete'")}) String sheetName,
            @Field.Text(name = "Column Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the column to be deleted'")}) String columnName) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Delete Column With Index", () -> {
            LOGGER.info("CatalogRequest deleteColumnWithIndex start with fileId: {}, driveId: {}, sheetName: {}, columnName: {}",
                    safeString(fileId), safeString(driveId), safeString(sheetName), safeString(columnName));
            return sharedOperationsFacade.getDeleteColumnWithIndexExtensionResponse(driveId, fileId, sheetName, columnName, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_48d01eb0-db22-40ca-8736-21a352489701",
            name = "Add Row With Index",
            description = "Insert Row At The Desired Index",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    public ExtensionResponse addRowWithIndex(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet where the row will be inserted'")}) String sheetName,
            @Field.Desc(name = "Rows", type = "[ Text ]", required = false) List<String> rows,
            @Field(name = "Row Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The index position where the row will be inserted'")}) Double rowIndex,
            @Field.File(name = "Rows From Csv", multipleFileUpload = true, required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'CSV file containing rows to be inserted at the specified index'")}) List<File> rowsFromCsv) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Add Row With Index", () -> {
            LOGGER.info("CatalogRequest addRowWithIndex start with driveId: {}, fileId: {}, sheetName: {}, rowIndex: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), rowIndex);
            return sharedOperationsFacade.getAddRowWithIndexExtensionResponse(driveId, fileId, sheetName, rows, rowsFromCsv, rowIndex, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_469adc90-8b8c-4b5e-a7b8-99b2bc1a36b1",
            name = "Add Row With Index Async",
            description = "Add a row with the specified index by passing the necessary information, as this is an async request.",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addRowWithIndexAsync(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String driveId,
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

                    restClientProvider.getRestClient(driveId).post(fileId +
                                    Constants.WORKBOOK_WORKSHEETS + sheetName + Constants.RANGE_ADDRESS
                                    + stringUtil.calculateAddressForInsert(rowIndex, rows,
                                    rowsFromCsv) + "')/insert",
                            Map.of(Constants.SHIFT, Constants.DOWN));
                    sheetFormattingService.clearRowFormatting(driveId, fileId, sheetName, rowIndex);

                    List<FreeForm> result = worksheetService.addRows(driveId, fileId, sheetName, rowsStringOrCsvRows,
                            StringUtil.calculateAddressForInsert(rowIndex, rowsStringOrCsvRows));

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Sheet Data", "[ Composite ]", result);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (IOException cause) {
                    LOGGER.error("Error in addRowWithIndexAsync with driveId: {}, fileId: {}, sheetName: {}, rowIndex: {}: {}",
                            driveId, fileId, sheetName, rowIndex, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_48d01eb0-db22-40ca-8736-21a352489701_result",
            name = "Get Add Row With Index Result",
            description = "Get the result of the asynchronous Add Row With Index operation.",
            area = "Shared",
            type = CatalogRequest.Type.WAIT_FOR_EVENT)
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    public List<FreeForm> getAddRowWithIndexResult(
            @Field(name = "eventName", type = "Text") String eventName,
            @Field(name = "eventData", type = "FreeForm") FreeForm eventData,
            @Field(name = "taskID", type = "Text") String taskID) {

        return audited("Get Add Row With Index Result", () -> {
            if (eventName.equals(taskID)) {
                return (List<FreeForm>) eventData.get("Sheet Data");
            }
            throw new IllegalStateException("Invalid task ID");
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_cdd0eb60-1f5a-4326-bdb2-fec6082f1148",
            name = "Update Row With Index",
            description = "Update Row With Index",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Desc(name = "Sheet Data", type = "[ Composite ]", required = false)
    public ExtensionResponse updateRowWithIndex(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive containing the Excel file'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet where the row will be updated'")}) String sheetName,
            @Field.Text(name = "Row", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The data to update the row with (comma-separated values)'")}) String row,
            @Field(name = "Row Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The index of the row to update'")}) Double rowIndex) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Update Row With Index", () -> {
            LOGGER.info("CatalogRequest updateRowWithIndex start with driveId: {}, fileId: {}, sheetName: {}, row: {}, rowIndex: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(row), rowIndex);
            return sharedOperationsFacade.getUpdateRowWithIndexExtensionResponse(driveId, fileId, sheetName, row, rowIndex, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_9f54eac1-2762-46c6-b568-af71eb66886f",
            name = "Get complete row info",
            description = "Based on the index should be able to get the complete row data",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field(name = "Row Data", type = "FreeForm", required = false)
    public ExtensionResponse getCompleteRowInfo(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the row to retrieve'")}) String sheetName,
            @Field(name = "Row Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The index of the row to retrieve data from'")}) Double rowIndex) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Get complete row info", () -> {
            LOGGER.info("CatalogRequest getCompleteRowInfo start with driveId: {}, fileId: {}, sheetName: {}, rowIndex: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), rowIndex);
            return sharedOperationsFacade.getCompleteRowInfoExtensionResponse(driveId, fileId, sheetName, rowIndex, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_4bbf5ef7-1763-459e-81c9-e76bee7aeafc",
            name = "Delete Row With Index",
            description = "Delete Row With Index",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Boolean(name = "Deleted", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public ExtensionResponse deleteRowWithIndex(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File ID", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileID,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the row to delete'")}) String sheetName,
            @Field(name = "Row Index", type = "Number", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The index of the row to be deleted'")}) Double rowIndex) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Delete Row With Index", () -> {
            LOGGER.info("CatalogRequest deleteRowWithIndex start with driveId: {}, fileId: {}, sheetName: {}, rowIndex: {}",
                    safeString(driveId), safeString(fileID), safeString(sheetName), rowIndex);
            return sharedOperationsFacade.getDeleteRowWithIndexExtensionResponse(driveId, fileID, sheetName, rowIndex, null);
        });
    }


    @CatalogRequest(
            id = "localDomainRequest_9c6e04d2-bba6-4054-b424-acca35b25a07",
            name = "Add Multiple Cell Data",
            description = "Add Multiple Cell Data",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Cell Data", type = "[ Composite ]", required = false)
    public ExtensionResponse addMultipleCellData(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String sheetName,
            @Field.Desc(name = "Cell ID Data", type = "[ { Cell Id: Text, Data: Text } ]") List<Map<String, Object>> cellIDData) {
        return audited("Add Multiple Cell Data", () -> {
            LOGGER.info("CatalogRequest addMultipleCellData start with driveId: {}, fileId: {}, sheetName: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName));
            return sharedOperationsFacade.getAddMultipleCellDataExtensionResponse(driveId, fileId, sheetName, cellIDData, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_6fd32754-025a-4556-ace0-e811288f114d",
            name = "Add Multiple Cell Data Async",
            description = "Add Multiple Cell Data Async",
            area = "Shared",
            type = CatalogRequest.Type.QUERY_SYSTEM)
    @Field.Text(name = "Task ID", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public String addMultipleCellDataAsync(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileId,
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
                        allCellData.add(this.worksheetService.addDataToCell(driveId, fileId, sheetName
                                , cellId, dataToBeAdded));
                    }

                    FreeForm freeForm = new FreeForm();
                    freeForm.put("Sheet Data", "[ Composite ]", allCellData);
                    eventHandler.handleEvent(taskId, freeForm);
                } catch (Exception cause) {
                    LOGGER.error("Error in addMultipleCellDataAsync with driveId: {}, fileId: {}, sheetName: {}: {}",
                            driveId, fileId, sheetName, cause.getMessage(), cause);
                    throw new RuntimeException(cause);
                } finally {
                    ThreadLocalProxy.removeAll();
                }
            });

            return taskId;
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_322f4be5-3b1d-4358-920a-ab2a092ff641",
            name = "Get Column Values For Keys In Sheet",
            description = "Get Column Values For Keys In Sheet",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Data", type = "[ Composite ]", required = false)
    public ExtensionResponse getColumnValuesForKeysInSheet(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the drive where the Excel file is located'")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The ID of the Excel file containing the sheet'")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The name of the sheet containing the data'")}) String sheetName,
            @Field.Text(name = "Range", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The range of cells to search in (e.g., A1:D10)'")}) String range,
            @Field.Desc(name = "Lookup Rows", type = "[ Text ]") List<String> lookupRows,
            @Field.Desc(name = "Lookup Columns", type = "[ Text ]") List<String> lookupColumns,
            @Field.Boolean(name = "Get All Records", attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'Whether to retrieve all matching records or just the first match'")}) Boolean getAllRecords,
            @Field(name = "Column Number", type = "Number", required = false, attributes = {@Attribute(name = "visualWidth", value = "S"), @Attribute(name = "toolTip", value = "'The column number to use as the key column (1-based index)'")}) Double columnNumber) {
        return audited("Get Column Values For Keys In Sheet", () -> {
            LOGGER.info("CatalogRequest getColumnValuesForKeysInSheet start with driveId: {}, fileId: {}, sheetName: {}, range: {}, lookupRows: {}, lookupColumns: {}, getAllRecords: {}, columnNumber: {}",
                    safeString(driveId), safeString(fileId), safeString(sheetName), safeString(range), lookupRows, lookupColumns, getAllRecords, columnNumber);
            return sharedOperationsFacade.getColumnValuesForKeysInSheetExtensionResponse(driveId, fileId, sheetName, range, lookupRows, lookupColumns, getAllRecords, columnNumber, null);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_bf7b30ea-e4f9-463e-9132-381ced204558",
            name = "Add Rows With List Of Entity",
            description = "Add Rows With List Of Entity",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM)
    @Field.Desc(name = "Total Rows", type = "[ Composite ]", required = false)
    public List<FreeForm> addRowsWithListOfEntity(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String driveId,
            @Field.Text(name = "File Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileId,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String sheetName,
            @Field.Desc(name = "Data", type = "[ Entity ]") List<EntityValue> data,
            @Field.Desc(name = "Exclude Columns", type = "[ Text ]", required = false) List<String> excludeColumns) {
        return audited("Add Rows With List Of Entity", () -> {
            if (StringUtil.isDataNullOrEmpty(data)) {
                return List.of((FreeForm) Constants.FREE_FORM.put("Error", "Text", "Data cannot be null or empty"));
            }
            EntityDefinition entityDefinition = registry.getEntityDefinition(data.getFirst());
            String usedRangeForSheet = worksheetOperation.getUsedRangeForSheet(driveId, fileId, sheetName);
            return this.worksheetOperation.addRowsToSheetWithEntity(fileId, sheetName, data, usedRangeForSheet, driveId, excludeColumns, entityDefinition);
        });
    }

    @CatalogRequest(
            id = "localDomainRequest_01b4ce79-ac11-4182-b56e-40d001b9613e",
            name = "Clear Shared Sheet For Range",
            description = "Clears the specified cell range in a shared sheet.",
            area = "Shared",
            type = CatalogRequest.Type.CHANGE_SYSTEM,
            tool = true)
    @Field.Boolean(name = "Cleared", required = false, attributes = {@Attribute(name = "visualWidth", value = "S")})
    public ExtensionResponse clearSharedSheetForRange(
            @Field.Text(name = "Drive Id", attributes = {@Attribute(name = "visualWidth", value = "S")}) String driveId,
            @Field.Text(name = "File ID", attributes = {@Attribute(name = "visualWidth", value = "S")}) String fileID,
            @Field.Text(name = "Sheet Name", attributes = {@Attribute(name = "visualWidth", value = "S")}) String sheetName,
            @Field.Text(name = "Cell Range", attributes = {@Attribute(name = "visualWidth", value = "S")}) String cellRange) {

        boolean useEmail = !requestContext.invokeAsUser();
        if (!useEmail) {
            providerFactory.create().validateToken(authorizationContext.getAuthorizedAccount().getAccountId());
        }
        return audited("Clear Shared Sheet For Range", () -> {
            LOGGER.info("CatalogRequest clearSharedSheetForRange start with driveId: {}, fileId: {}, sheetName: {}, cellRange: {}",
                    safeString(driveId), safeString(fileID), safeString(sheetName), safeString(cellRange));
            return sharedOperationsFacade.getClearSharedSheetForRangeExtensionResponse(driveId, fileID, sheetName, cellRange, null);
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
