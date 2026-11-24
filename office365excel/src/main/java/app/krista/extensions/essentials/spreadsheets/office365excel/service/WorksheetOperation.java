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

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.DataModification;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Tables;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.parser.DataParser;
import app.krista.model.base.EntityValue;
import app.krista.model.base.FreeForm;
import app.krista.model.entity.EntityDefinition;
import com.google.gson.JsonElement;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.requests.WorkbookWorksheetRequestBuilder;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

@Service
public class WorksheetOperation {
    private final GraphServiceClientProviderFactory providerFactory;
    private GraphServiceClientProvider clientProvider;
    private final RestClientProvider restClientProvider;
    private final WorkbookOperation workbookOperation;
    public static final String WORKBOOK_WORKSHEETS = "/workbook/worksheets/";
    public static final String RANGE_ADDRESS = "/range(address='";
    public static final String STRING = "')";
    public static final String ERROR = "error";
    public static final String TEXT = "Text";

    private static final Logger LOGGER = LoggerFactory.getLogger(WorksheetOperation.class);

    @Inject
    public WorksheetOperation(WorkbookOperation workbookOperation, GraphServiceClientProviderFactory providerFactory, RestClientProvider restClientProvider) {
        this.workbookOperation = workbookOperation;
        this.providerFactory = providerFactory;
        this.restClientProvider = restClientProvider;
    }

    /** Lazy provider creation — avoids MustAuthorizeException during HK2 construction. */
    private GraphServiceClientProvider getClientProvider() {
        if (clientProvider == null) {
            clientProvider = providerFactory.create();
        }
        return clientProvider;
    }

    public List<FreeForm> getRowAndColumnData(String workbookID,
                                              String tableName, List<String> keyRows, List<String> lookupColumns, Boolean getAllRecord,
                                              Double columnNumber, String sheetName, String range, String driveId) {
        JsonElement text;
        keyRows = uniqueKeyRows(keyRows);

        try {
            String string = null;
            WorkbookRange workbookRange = null;
            if (tableName != null) {
                workbookRange = getWorkbookRange(driveId,
                        workbookID, tableName);
            } else {
                string = Objects.requireNonNull(restClientProvider.getRestClient(driveId).get(
                        workbookID + WORKBOOK_WORKSHEETS + sheetName + RANGE_ADDRESS + range
                                + STRING).body()).string();
            }
            if (workbookRange != null) {
                text = workbookRange.text;
            } else if (string != null) {
                text = Constants.GSON.fromJson(string, JsonElement.class);
            } else {
                return List.of(new FreeForm());
            }

            List<FreeForm> errorMsgForText = errorMessageForText(text);
            if (errorMsgForText != null) {
                return errorMsgForText;
            }

            List<String> columnHeaders = DataParser.getColumnHeaders(text);
            List<FreeForm> errorMsg = errorMessage(columnNumber, columnHeaders);
            if (errorMsg != null) {
                return errorMsg;
            }

            String keyColumnHeader = getColumnHeader(columnNumber, columnHeaders);

            List<FreeForm> rowDataList = DataParser.parseTableDataForColumnValue(text);

            return Tables.getRowsDataForColumns(rowDataList, keyColumnHeader, keyRows, lookupColumns, getAllRecord);
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to get row and column data: {}", cause.getMessage(), cause);
            return Collections.emptyList();
        } catch (IOException cause) {
            throw new RuntimeException(cause);
        }
    }

    private static List<FreeForm> errorMessageForText(JsonElement text) {
        if (text == null) {
            FreeForm freeForm = Constants.FREE_FORM;
            freeForm.put(ERROR, TEXT, "There is no data available for parsing");
            return List.of(freeForm);
        }
        return null;
    }

    private static List<FreeForm> errorMessage(Double columnNumber, List<String> columnHeaders) {
        FreeForm freeForm = Constants.FREE_FORM;
        if (columnNumber == null) {
            return null; // should return null based on logic implemented and not empty list
        }
        if (columnHeaders == null || columnHeaders.isEmpty()) {
            freeForm.put(ERROR, TEXT, "ColumnHeader should not be empty or null");
            return List.of(freeForm);
        }
        if (columnNumber - 1 >= columnHeaders.size() || columnNumber <= 0) {

            freeForm.put(ERROR, TEXT, "Column number " + columnNumber.intValue() + " should be less or equal to  " + columnHeaders.size() + " and should not be less than 1");
            return List.of(freeForm);
        }
        return null;
    }

    private List<String> uniqueKeyRows(List<String> keyRows) {
        return keyRows.stream().distinct().collect(Collectors.toList());
    }

    private static String getColumnHeader(Double columnNumber, List<String> columnHeaders) {
        if (columnNumber == null) {
            return columnHeaders.getFirst();
        }
        return columnNumber - 1 <= columnHeaders.size()
                ? columnHeaders.get(columnNumber.intValue() - 1)
                : columnHeaders.getFirst();
    }

    private WorkbookRange getWorkbookRange(String driveId,
                                           String workbookID, String tableName) {

        return getClientProvider().getDriveRequestBuilder(driveId).items(workbookID).workbook().tables(tableName).range()
                .buildRequest().get();
    }

    public String getUsedRangeForSheet(String driveId, String workbookId,
                                       String worksheetName) {
        WorkbookWorksheetRequestBuilder workbookWorksheetRequestBuilder = getClientProvider().worksheetRequestBuilders(
                driveId, worksheetName, workbookId);
        return Objects.requireNonNull(
                workbookWorksheetRequestBuilder.usedRange().buildRequest().get()).address;
    }

    public List<FreeForm> addRowsToSheetWithEntity(String fileID, String sheetName, List<EntityValue> data, String usedRangeForSheet, String driveId, List<String> excludeColumns, EntityDefinition entityDefinition) {
        List<List<String>> lists = DataModification.convertToListOfListsFromIndex(
                DataModification.convertToListOfLists(DataModification.fromListOfEntities(data, excludeColumns, entityDefinition, getClientProvider())),
                usedRangeForSheet.contains(Constants.COLON) ? 1 : 0);

        return workbookOperation.addRowsForListOfEntity(driveId, fileID, sheetName, lists);
    }

}
