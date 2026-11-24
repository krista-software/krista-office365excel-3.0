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

package app.krista.extensions.essentials.spreadsheets.office365excel.service;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelSheet;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Numbers;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Tables;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.parser.DataParser;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonElement;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.*;
import com.microsoft.graph.requests.WorkbookRequestBuilder;
import com.microsoft.graph.requests.WorkbookTableCollectionPage;
import com.microsoft.graph.requests.WorkbookTableCollectionRequestBuilder;
import com.microsoft.graph.requests.WorkbookTableRangeRequestBuilder;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.List;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class TableService {
    private final WorksheetService worksheetService;
    private final GraphServiceClientProviderFactory providerFactory;
    private GraphServiceClientProvider clientProvider;
    private static final Logger LOGGER = LoggerFactory.getLogger(TableService.class);


    @Inject
    public TableService(WorksheetService worksheetService, GraphServiceClientProviderFactory providerFactory) {
        this.worksheetService = worksheetService;
        this.providerFactory = providerFactory;
    }

    /** Lazy provider creation — avoids MustAuthorizeException during HK2 construction. */
    private GraphServiceClientProvider getClientProvider() {
        if (clientProvider == null) {
            clientProvider = providerFactory.create();
        }
        return clientProvider;
    }

    /**
     * Creates a new table in an Excel worksheet at the specified range.
     * The range must include headers in the first row.
     *
     * @param workbookId the ID of the Excel workbook
     * @param worksheetName the name of the worksheet where the table will be created
     * @param range the cell range for the table including headers (e.g., "A1:D10")
     * @param driveId the ID of the drive containing the workbook
     * @return Table object containing the created table information
     * @throws IllegalArgumentException if table creation fails due to invalid range, missing worksheet, or permissions
     */
    public Table createWorkbookTable(String workbookId, String worksheetName, String range,
                                     String driveId) {
        try {
            WorkbookTableAddParameterSet tableAddParameterSet = WorkbookTableAddParameterSet
                    .newBuilder()
                    .withAddress(range)
                    .build();

            WorkbookTable workbookTable = clientProvider
                    .getDriveRequestBuilder(driveId)
                    .items(workbookId)
                    .workbook()
                    .worksheets(worksheetName)
                    .tables()
                    .add(tableAddParameterSet)
                    .buildRequest()
                    .post();

            return DataParser.parse(workbookTable, worksheetName);
        } catch (GraphServiceException cause) {
            String errorMessage = cause.getServiceError() != null ? cause.getServiceError().message : cause.getMessage();
            LOGGER.error("Failed to create table: {}", errorMessage, cause);
            throw new IllegalArgumentException("Unable to create table in Excel file. " + errorMessage + ". " +
                    "Please verify that the worksheet exists, the range is valid, and you have permission to modify the file.");
        }
    }

    /**
     * Retrieves all data from an Excel table including headers and rows.
     * Returns data as a list of FreeForm objects where each object represents a row.
     *
     * @param driveId the ID of the drive containing the workbook
     * @param workbookID the ID of the Excel workbook
     * @param tableName the name of the table to retrieve data from
     * @return list of FreeForm objects representing table data, empty list if table not found or on error
     */
    public List<FreeForm> getTableData(String driveId, String workbookID, String tableName) {
        try {
            WorkbookTableRangeRequestBuilder tableRangeRequestBuilder = getClientProvider().getDriveRequestBuilder(driveId)
                    .items(workbookID)
                    .workbook()
                    .tables(tableName)
                    .range();

            WorkbookRange tableRange = tableRangeRequestBuilder.buildRequest().get();
            if (tableRange == null) {
                LOGGER.error("Table range is null for table: {}", tableName);
                return Collections.emptyList();
            }
            JsonElement jsonElement = tableRange.text;
            return DataParser.parseTableData(jsonElement);
        } catch (GraphServiceException cause) {
            String errorMessage = cause.getServiceError() != null ? cause.getServiceError().message : cause.getMessage();
            LOGGER.error("Failed to get table data for table '{}': {}", tableName, errorMessage, cause);
            return Collections.emptyList();
        }
    }

    /**
     * This function gives the tableNames along with information like in  which sheet it belongs to .
     *
     * @param accountId    for the verification process need to pass the accountId of the User
     * @param invokeAsUser Tests if the request is invoked as user
     * @param fileId       {@link String} refers to the Excel File ID
     * @param driveId      {@link String} refers to the Excel Drive ID
     * @return {@link List<Table>} gives the details about the location of table,  that is  in which sheet it belongs
     */
    public List<Table> getTableNames(String accountId, boolean invokeAsUser, String fileId, String driveId) {
        List<Table> allTheTable = new ArrayList<>();
        List<WorkbookTable> workbookTables = new ArrayList<>();
        List<String> sheetNames = getSheetNames(fileId, driveId, accountId, invokeAsUser);
        Map<String, WorkbookTableCollectionPage> workbookTableCollectionPageMap = listOfTables(fileId, driveId, sheetNames, accountId, invokeAsUser);
        if (workbookTableCollectionPageMap == null) {
            return List.of();
        }

        for (Map.Entry<String, WorkbookTableCollectionPage> pages : workbookTableCollectionPageMap.entrySet()) {
            WorkbookTableCollectionPage eachWorkbookTableCollectionPage = pages.getValue();
            if (eachWorkbookTableCollectionPage == null) {
                return List.of();
            }
            String SheetName = pages.getKey();
            while (eachWorkbookTableCollectionPage != null && !eachWorkbookTableCollectionPage.getCurrentPage().isEmpty()) {
                workbookTables.addAll(eachWorkbookTableCollectionPage.getCurrentPage());
                WorkbookTableCollectionRequestBuilder nextPage = eachWorkbookTableCollectionPage.getNextPage();
                if (nextPage != null) {
                    eachWorkbookTableCollectionPage = nextPage.buildRequest().get();
                } else {
                    break;
                }

            }
            allTheTable.addAll(Tables.getTables(workbookTables, SheetName));
            workbookTables.clear();
        }


        return allTheTable;
    }

    /**
     * @param fileID       {@link String} refers to the Excel File ID
     * @param driveID      {@link String} refers to the Excel Drive ID
     * @param accountId    for the verification process need to pass the accountId of the User
     * @param invokeAsUser Tests if the request is invoked as user
     * @return {@link List<String>} return the all the sheetNames present in the Excel File
     */
    private List<String> getSheetNames(String fileID, String driveID, String accountId, boolean invokeAsUser) {
        return this.worksheetService.getWorksheets(fileID, driveID, accountId, invokeAsUser).stream().map(ExcelSheet::getSheetName).collect(Collectors.toList());
    }

    private Map<String, WorkbookTableCollectionPage> listOfTables(String workbookId, String DriveId, List<String> sheetNames, String accountId, boolean invokeAsUser) {
        Map<String, WorkbookTableCollectionPage> workbookTableCollectionPagesCollection = new LinkedHashMap<>();
        for (String names : sheetNames) {
            try {
                WorkbookTableCollectionPage eachPage = getClientProvider().getDriveRequestBuilder(DriveId, accountId, invokeAsUser).
                        items(workbookId).
                        workbook().
                        worksheets(names).
                        tables().
                        buildRequest().
                        get();
                workbookTableCollectionPagesCollection.put(names, eachPage);

            } catch (GraphServiceException | IOException cause) {
                LOGGER.error("Failed to get list of tables: {}", cause.getMessage(), cause);
                return null;
            }
        }
        return workbookTableCollectionPagesCollection;
    }

    public FreeForm vlookup(String driveId, String fileId, String lookupValue,
                            String tableName, Double columnIndex, Boolean exactMatch) {
        try {
            JsonElement lookupVal = Constants.GSON.toJsonTree(Numbers.isNumber(lookupValue) ? Double.parseDouble(lookupValue) : lookupValue);
            JsonElement colIndex = Constants.GSON.toJsonTree(columnIndex);
            JsonElement rangeLookup = Constants.GSON.toJsonTree(exactMatch);

            WorkbookRequestBuilder workbook = getClientProvider().getDriveRequestBuilder(driveId).items(fileId).workbook();
            WorkbookTableRangeRequestBuilder rangeBuilder = workbook.tables(tableName.trim()).range();
            if (rangeBuilder.buildRequest().get() == null) {
                LOGGER.error("Range builder is null");
                return new FreeForm();
            }
            JsonElement rangeJsonElem = Objects.requireNonNull(rangeBuilder.buildRequest().get()).values;

            WorkbookFunctionsVlookupParameterSet vLookupParameterSet = WorkbookFunctionsVlookupParameterSet
                    .newBuilder()
                    .withLookupValue(lookupVal)
                    .withRangeLookup(rangeLookup)
                    .withColIndexNum(colIndex)
                    .withTableArray(rangeJsonElem)
                    .build();

            WorkbookFunctionResult functionResult = workbook.functions().vlookup(vLookupParameterSet).buildRequest().post();
            if (functionResult == null) {
                LOGGER.error("Function result is null");
                return new FreeForm();
            }
            JsonElement element = functionResult.value;
            if (element == null) {
                LOGGER.error("Element is null");
                return new FreeForm();
            }
            Object value = DataParser.getCellData(element);
            FreeForm freeForm = new FreeForm();
            freeForm.put(lookupValue, "Text", value);
            return freeForm;
        } catch (GraphServiceException cause) {
            LOGGER.error("Cause:{}", cause.getMessage(), cause);
            return new FreeForm();
        }
    }
}
