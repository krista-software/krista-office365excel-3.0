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

import app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2.WorkbookWorksheetRangeRequestBuilderV2;
import app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2.WorkbookWorksheetRangeRequestV2;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Workbooks;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.parser.DataParser;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.DriveSearchParameterSet;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.models.WorkbookWorksheetRangeParameterSet;
import com.microsoft.graph.requests.DriveSearchCollectionPage;
import com.microsoft.graph.requests.WorkbookWorksheetRangeRequestBuilder;
import com.microsoft.graph.requests.WorkbookWorksheetRequestBuilder;
import com.microsoft.graph.requests.WorkbookWorksheetUsedRangeRequest;
import okhttp3.ResponseBody;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Service
public class WorkbookOperation {

    public static final String UPDATED_WORK_BOOK_RANGE_ADDRESS = "updatedWorkBookRangeAddress = ";
    public static final String PROBLEM_WITH_WORKBOOK_RANGE = "Problem with workbookRange";
    public static final String SHEET_IS_EMPTY = "Sheet Is Empty";
    public static final String RESPONSE_BODY_IS_EMPTY = "Response Body Is Empty";
    public static final String RANGE_ADDRESS = "/range(address='";
    public static final String WORKBOOK_WORKSHEETS_SHEET = "/workbook/worksheets/";
    private GraphServiceClientProvider clientProvider;
    private final GraphServiceClientProviderFactory providerFactory;
    private final RestClientProvider restClientProvider;
    private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookOperation.class);


    @Inject
    public WorkbookOperation(GraphServiceClientProviderFactory providerFactory, RestClientProvider restClientProvider) {
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

    /**
     * Searches for items in the drive collection by name.
     * Uses Microsoft Graph API search functionality to find files and folders.
     *
     * @param itemName the name or partial name of the item to search for
     * @return DriveSearchCollectionPage containing search results, null on error
     */
    public DriveSearchCollectionPage searchDriveCollection(String itemName) {
        try {
            DriveSearchParameterSet searchParameterSet = DriveSearchParameterSet.newBuilder().withQ(itemName).build();
            return getClientProvider().getDriveRequestBuilder(null)
                    .search(searchParameterSet).buildRequest().get();
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to search drive collection: {}", cause.getMessage(), cause);
            return null;
        }
    }

    /**
     * Adds rows to an Excel worksheet at a specified address.
     * Retrieves the current used range, updates it with new rows, and patches the worksheet.
     * Handles blank sheets by returning empty list instead of a single blank row.
     *
     * @param driveId the ID of the drive containing the workbook
     * @param workbookId the ID of the Excel workbook
     * @param worksheetName the name of the worksheet to add rows to
     * @param rows list of comma-separated row data strings
     * @param address the cell address where rows should be added (e.g., "A1")
     * @return list of FreeForm objects representing all data in the worksheet after addition, null on error
     */
    public List<FreeForm> addRows(String driveId, String workbookId, String worksheetName, List<String> rows, String address) {
        try {
            WorkbookWorksheetRequestBuilder worksheetRequestBuilder = getClientProvider().worksheetRequestBuilders(driveId, worksheetName, workbookId);
            WorkbookWorksheetUsedRangeRequest usedRangeRequest = worksheetRequestBuilder.usedRange().buildRequest();
            WorkbookRange workbookRange = usedRangeRequest.get();
            if (workbookRange == null) {
                LOGGER.error("Workbook range is null");
                return null;
            }
            WorkbookRange updatedWorkbookRange = Worksheets.updateWorkbookRange(worksheetName, workbookRange, rows, address);
            LOGGER.info(UPDATED_WORK_BOOK_RANGE_ADDRESS + "{}", updatedWorkbookRange.address);

            WorkbookWorksheetRangeParameterSet rangeParameterSet = Workbooks.getWorkbookWorksheetRangeParameterSet(Workbooks.extractRange(updatedWorkbookRange.address));
            WorkbookWorksheetRangeRequestBuilder rangeRequest = worksheetRequestBuilder.range(rangeParameterSet);

            if (rangeRequest.getRequestUrl() == null || rangeRequest.getClient() == null) {
                LOGGER.error("Range request is null");
                return null;
            }
            WorkbookWorksheetRangeRequestBuilderV2 rangeRequestBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    rangeRequest.getRequestUrl(), rangeRequest.getClient(), rangeRequest.getOptions(), rangeParameterSet);
            WorkbookWorksheetRangeRequestV2 rangeRequestV2 = rangeRequestBuilderV2.buildRequest();
            WorkbookRange wbRange = rangeRequestV2.patch(updatedWorkbookRange);
            if (wbRange == null || wbRange.values == null || workbookRange.values == null) {
                LOGGER.error("Workbook range is null");
                return null;
            }
            List<FreeForm> finalData = DataParser.parse(workbookRange.values);
            List<FreeForm> newData = DataParser.parse(wbRange.values);
            if (finalData.isEmpty()) {
                // required to avoid the first blank line in client side
                /* See if you go inside parse function , Api returns[[““]] if the sheet is blank and if this is the case ,
                 than it would make the size of array as 1 with blank row which is completely useless , instead just return
                 empty list
                */
                return newData;
            }

            finalData.addAll(Workbooks.getFreeForms(finalData, newData));
            return finalData;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to add rows: {}", cause.getMessage(), cause);
            return null;
        }

    }

    public List<FreeForm> addRowsForInsertAndUpdate(String driveId, String workbookId,
                                                    String worksheetName, List<String> rows, String address) {

        LOGGER.info("Starting addRowsForInsertAndUpdate - worksheet: {}, rows count: {}, address: {}",
                worksheetName, rows != null ? rows.size() : 0, address);

        if (rows != null && !rows.isEmpty()) {
            LOGGER.info("First row content: {}", rows.getFirst());

            boolean containsHyperlink = rows.stream().anyMatch(row -> row.contains("=HYPERLINK"));
            if (containsHyperlink) {
                LOGGER.info("Input contains HYPERLINK formulas - special handling will be applied");
            }
        }

        try {
            // Step 1: Get worksheet request builder
            WorkbookWorksheetRequestBuilder worksheetRequestBuilder = getClientProvider().worksheetRequestBuilders(
                    driveId, worksheetName, workbookId);
            LOGGER.debug("Created worksheet request builder");

            // Step 2: Get the current used range
            WorkbookWorksheetUsedRangeRequest usedRangeRequest = worksheetRequestBuilder.usedRange().buildRequest();
            WorkbookRange workbookRange = usedRangeRequest.get();
            if (workbookRange == null) {
                LOGGER.error("Workbook range is null");
                return null;
            }
            LOGGER.debug("Retrieved current used range: {}", workbookRange.address);

            // Step 3: Update the workbook range with new rows
            LOGGER.info("Updating workbook range with {} rows", rows == null ? 0 : rows.size());
            WorkbookRange updatedWorkbookRange = Worksheets.updateWorkbookRange(worksheetName,
                    workbookRange, rows, address);
            LOGGER.info(UPDATED_WORK_BOOK_RANGE_ADDRESS + "{}", updatedWorkbookRange.address);
            if (updatedWorkbookRange.address == null) {
                LOGGER.error("Updated workbook range address is null");
                return null;
            }
            // Step 4: Create range parameter set and request builder
            String extractedRange = Workbooks.extractRange(updatedWorkbookRange.address);
            LOGGER.debug("Extracted range: {}", extractedRange);
            WorkbookWorksheetRangeParameterSet rangeParameterSet =
                    Workbooks.getWorkbookWorksheetRangeParameterSet(extractedRange);

            WorkbookWorksheetRangeRequestBuilder rangeRequest = worksheetRequestBuilder.range(
                    rangeParameterSet);
            if (rangeRequest.getRequestUrl() == null || rangeRequest.getClient() == null) {
                LOGGER.error("Range request is null");
                return null;
            }
            // Step 5: Create V2 request builder and patch the range
            LOGGER.debug("Creating V2 request builder for patching");
            WorkbookWorksheetRangeRequestBuilderV2 rangeRequestBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    rangeRequest.getRequestUrl(), rangeRequest.getClient(), rangeRequest.getOptions(),
                    rangeParameterSet);

            WorkbookWorksheetRangeRequestV2 rangeRequestV2 = rangeRequestBuilderV2.buildRequest();

            // Check if we have formulas in the updated range
            boolean hasFormulas = updatedWorkbookRange.formulas != null;
            LOGGER.info("Range contains formulas: {}", hasFormulas);
            if (hasFormulas) {
                LOGGER.debug("Formula data: {}", updatedWorkbookRange.formulas);
            }

            // Patch the range
            LOGGER.info("Patching workbook range");
            rangeRequestV2.patch(updatedWorkbookRange);

            // Need some time between the patch and consecutive getUsedRange request to properly range used by the sheet after adding the row
            LOGGER.info("Waiting for Excel to process the changes...");
            Thread.sleep(5000);

            // Step 6: Get the updated used range
            WorkbookRange usedRange = getUsedRange(driveId, worksheetName, workbookId);
            if (usedRange == null || usedRange.address == null) {
                LOGGER.error("Used range is null");
                return null;
            }
            LOGGER.info("Retrieved updated used range: {}", usedRange.address);

            // Step 7: Get data for the range and return
            String finalRange = Workbooks.extractRange(usedRange.address);
            LOGGER.info("Getting data for range: {}", finalRange);
            List<FreeForm> result = getDataForRange(driveId, workbookId, worksheetName, finalRange);
            LOGGER.info("Retrieved {} rows of data", result == null ? 0 : result.size());

            return result;
        } catch (GraphServiceException cause) {
            LOGGER.error("GraphServiceException in addRowsForInsertAndUpdate: {}", cause.getMessage(), cause);
            return null;
        } catch (InterruptedException cause) {
            LOGGER.error("InterruptedException in addRowsForInsertAndUpdate: {}", cause.getMessage(), cause);
            Thread.currentThread().interrupt(); // Restore the interrupted status
            return null;
        } catch (Exception cause) {
            LOGGER.error("Exception in addRowsForInsertAndUpdate: {}", cause.getMessage(), cause);
            return null;
        }
    }

    private WorkbookRange getUsedRange(String driveId, String worksheetName
            , String workbookId) {

        WorkbookWorksheetRequestBuilder worksheetRequestBuilder = getClientProvider().worksheetRequestBuilders(
                driveId, worksheetName, workbookId);
        WorkbookWorksheetUsedRangeRequest usedRangeRequest = worksheetRequestBuilder.usedRange()
                .buildRequest();

        return usedRangeRequest.get();
    }

    private List<FreeForm> getDataForRange(String driveId, String workbookID,
                                           String worksheetName, String range) {
        WorkbookWorksheetRangeParameterSet parameterSet = Workbooks.getWorkbookWorksheetRangeParameterSet(range);

        WorkbookWorksheetRangeRequestBuilder baseBuilder = getClientProvider().worksheetRequestBuilders(driveId, worksheetName, workbookID)
                .range(parameterSet);
        if (baseBuilder.getRequestUrl() == null || baseBuilder.getClient() == null) {
            LOGGER.error("Base builder is null");
            return null;
        }
        WorkbookWorksheetRangeRequestBuilderV2 builderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                baseBuilder.getRequestUrl(), baseBuilder.getClient(), baseBuilder.getOptions(), parameterSet);
        WorkbookRange workbookRange = builderV2.buildRequest().get();
        if (workbookRange == null) {
            LOGGER.error("Workbook range is null");
            return null;
        }
        JsonElement jsonElement = workbookRange.text;
        if (jsonElement == null) {
            LOGGER.error("JSON element is null");
            return null;
        }
        return DataParser.parse(jsonElement);
    }

    public FreeForm getRow(String fileID, String sheetName, Double rowIndex, String driveId) {

        FreeForm finalRow = new FreeForm();
        List<String> dummyRows = new ArrayList<>();
        try {
            WorkbookRange workbookRange = getClientProvider().worksheetRequestBuilders(driveId, sheetName, fileID)
                    .usedRange()
                    .buildRequest()
                    .get();
            if (workbookRange == null) {
                finalRow.put(Constants.ERROR, Constants.TEXT, PROBLEM_WITH_WORKBOOK_RANGE);
                return finalRow;
            }

            String address = workbookRange.address;
            if (address != null && !StringUtil.checkColonPresence(address)) {
                return (FreeForm) new FreeForm().put(Constants.ERROR, Constants.TEXT,
                        SHEET_IS_EMPTY);
            } else {
                return getCompleteRowInfo(fileID, rowIndex, finalRow, dummyRows, address, driveId, sheetName);
            }
        } catch (IOException cause) {
            throw new RuntimeException(cause);
        }
    }

    private FreeForm getCompleteRowInfo(String fileID, Double rowIndex, FreeForm finalRow,
                                        List<String> dummyRows, String address, String driveId, String sheetName) throws IOException {

        String finalAddress = "A" + rowIndex.intValue() + ":" + StringUtil.extractAlphabetAfterColon(
                address) + rowIndex.intValue();

        ResponseBody body = restClientProvider.getRestClient(driveId)
                .get(fileID + WORKBOOK_WORKSHEETS_SHEET + sheetName + RANGE_ADDRESS + finalAddress
                        + "')").body();

        if (body == null) {
            finalRow.put(Constants.ERROR, Constants.TEXT, RESPONSE_BODY_IS_EMPTY);
            return finalRow;
        }
        JsonArray jsonArray = Constants.GSON.fromJson(body.string(), JsonObject.class).
                get("text").getAsJsonArray().get(0).getAsJsonArray();

        Workbooks.extractRequiredData(finalRow, dummyRows, jsonArray);
        return finalRow;
    }

    public List<FreeForm> addRowsForListOfEntity(String driveId, String workbookId,
                                                 String worksheetName, List<List<String>> rows) {
        try {
            WorkbookWorksheetRequestBuilder worksheetRequestBuilder = getClientProvider().worksheetRequestBuilders(
                    driveId, worksheetName, workbookId);
            WorkbookWorksheetUsedRangeRequest usedRangeRequest = worksheetRequestBuilder.usedRange()
                    .buildRequest();
            WorkbookRange workbookRange = usedRangeRequest.get();
            if (workbookRange == null) {
                LOGGER.error("Workbook range is null");
                return null;
            }
            WorkbookRange updatedWorkbookRange = Worksheets.updateWorkbookRangeForEntity(worksheetName,
                    workbookRange, rows, null);
            if (updatedWorkbookRange.address == null) {
                LOGGER.error("Updated workbook range is null");
                return null;
            }
            LOGGER.info(UPDATED_WORK_BOOK_RANGE_ADDRESS + "{}", updatedWorkbookRange.address);
            WorkbookWorksheetRangeParameterSet rangeParameterSet = Workbooks.getWorkbookWorksheetRangeParameterSet(Workbooks.extractRange(updatedWorkbookRange.address));
            WorkbookWorksheetRangeRequestBuilder rangeRequest = worksheetRequestBuilder.range(
                    rangeParameterSet);
            WorkbookWorksheetRangeRequestBuilderV2 rangeRequestBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    Objects.requireNonNull(rangeRequest.getRequestUrl()), Objects.requireNonNull(rangeRequest.getClient()), rangeRequest.getOptions(),
                    rangeParameterSet);
            WorkbookWorksheetRangeRequestV2 rangeRequestV2 = rangeRequestBuilderV2.buildRequest();
            rangeRequestV2.patch(updatedWorkbookRange);
              /*
              Need some time between the patch and consecutive getUsedRange request to properly
              range used by the sheet after adding the row
              */
            Thread.sleep(3000);
            WorkbookRange usedRange = getUsedRange(driveId, worksheetName, workbookId);
            if (usedRange.address == null) {
                LOGGER.error("Used range address is null");
                return null;
            }
            LOGGER.info("Used Range After Patch Request {}", usedRange.address);
            return getDataForRange(driveId, workbookId, worksheetName, Workbooks.extractRange(usedRange.address));
        } catch (GraphServiceException | InterruptedException cause) {
            LOGGER.debug("addRowsForListOfEntity method gives following error{}", cause.getMessage(), cause);
            return List.of();
        }
    }

}
