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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2.WorkbookWorksheetRangeRequestBuilderV2;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Tables;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Workbooks;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.parser.DataParser;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.File;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonElement;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.*;
import com.microsoft.graph.requests.*;
import org.jetbrains.annotations.NotNull;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.GSON;

@Service
public class WorkbookService {

    private final WorkbookOperation workbookOperation;
    private final GraphServiceClientProviderFactory providerFactory;
    private GraphServiceClientProvider clientProvider;
    private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookService.class);
    private final KristaMediaClient kristaMediaClient;

    /**
     * Context class to hold batch processing state and configuration.
     */
    private static class BatchProcessingContext {
        final String driveId;
        final String fileId;
        final String tableName;
        final List<String> allRows;
        final int columnCount;
        final int batchSize;
        final int totalRows;
        final long startTime;
        final app.krista.extensions.util.EventHandler eventHandler;
        final String taskId;
        final int[] progressMilestones = {20, 40, 60, 80}; // 100% handled separately in completion

        int currentMilestoneIndex = 0;
        int processedRows = 0;
        final List<FreeForm> allResults = new ArrayList<>();

        BatchProcessingContext(String driveId, String fileId, String tableName, List<String> allRows,
                               int columnCount, app.krista.extensions.util.EventHandler eventHandler, String taskId) {
            this.driveId = driveId;
            this.fileId = fileId;
            this.tableName = tableName;
            this.allRows = allRows;
            this.columnCount = columnCount;
            this.totalRows = allRows.size();

            // Calculate batch size based on total rows
            if (allRows.size() >= 2000) {
                // Use the standard calculation if batch size is >= 2000
                this.batchSize = 15000 / columnCount;
            } else {
                // For smaller datasets (< 2000), calculate batch size based on total rows
                // to ensure proper progress alignment
                this.batchSize = Math.max(1, this.totalRows / 5); // Divide into ~5 batches for better progress tracking
            }

            this.startTime = System.currentTimeMillis();
            this.eventHandler = eventHandler;
            this.taskId = taskId;
        }
    }

    private final TableService tableService;


    @Inject
    public WorkbookService(WorkbookOperation workbookOperation, GraphServiceClientProviderFactory providerFactory, KristaMediaClient kristaMediaClient, TableService tableService) {
        this.providerFactory = providerFactory;
        this.workbookOperation = workbookOperation;
        this.kristaMediaClient = kristaMediaClient;
        this.tableService = tableService;
    }

    public WorkbookService(WorkbookOperation workbookOperation, GraphServiceClientProvider provider, KristaMediaClient kristaMediaClient, TableService tableService) {
        this.providerFactory = null;
        this.clientProvider = provider;
        this.workbookOperation = workbookOperation;
        this.kristaMediaClient = kristaMediaClient;
        this.tableService = tableService;
    }

    /** Lazy provider creation — avoids MustAuthorizeException during HK2 construction. */
    private GraphServiceClientProvider getClientProvider() {
        if (clientProvider == null) {
            clientProvider = providerFactory.create();
        }
        return clientProvider;
    }

    /**
     * Retrieves detailed information about an Excel workbook from OneDrive.
     *
     * @param workbookFolderPath the folder path where the workbook is located (e.g., "/Documents")
     * @param workbookName the name of the workbook file (e.g., "MyWorkbook.xlsx")
     * @return ExcelFile object containing workbook ID, drive ID, name, and web URL
     * @throws IllegalArgumentException if the workbook is not found at the specified path
     * @throws RuntimeException if a Graph API error occurs during retrieval
     */
    public ExcelFile getWorkbookDetails(String workbookFolderPath, String workbookName) {
        try {
            String workbookPath = Workbooks.getPath(workbookFolderPath, workbookName);
            DriveItem item = getClientProvider().getDriveRequestBuilder(null).root().itemWithPath(workbookPath).buildRequest().get();
            if (item == null || item.parentReference == null) {
                throw new IllegalArgumentException("Excel workbook not found at the specified path. Please verify the file exists and you have permission to access it.");
            }
            return new ExcelFile(item.id, item.parentReference.driveId, item.name, item.webUrl);
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to get workbook details: {}", cause.getMessage(), cause);
            throw new RuntimeException(cause);
        }
    }

    /**
     * Searches for Excel workbooks by name across all accessible drives.
     * Paginates through all search results to return a complete list of matching workbooks.
     *
     * @param workbookName the name or partial name of the workbook to search for
     * @return list of ExcelFile objects matching the search criteria, empty list if none found
     */
    public List<ExcelFile> getWorkbooks(String workbookName) {

        DriveSearchCollectionPage collectionPage = workbookOperation.searchDriveCollection(workbookName);
        List<DriveItem> driveItems = new ArrayList<>();
        while (collectionPage != null) {
            driveItems.addAll(collectionPage.getCurrentPage());
            DriveSearchCollectionRequestBuilder nextPage = collectionPage.getNextPage();
            if (nextPage != null) {
                collectionPage = nextPage.buildRequest().get();
            } else {
                break;
            }
        }
        return Workbooks.getWorkbooks(workbookName, driveItems);
    }

    /**
     * Adds rows to an Excel table in batches with progress tracking and retry logic.
     * Implements dynamic batch sizing based on dataset size and column count.
     * Provides progress notifications at 20%, 40%, 60%, 80%, and 100% completion.
     *
     * @param driveId the ID of the drive containing the workbook
     * @param fileId the ID of the Excel file
     * @param tableName the name of the table to add rows to
     * @param rows list of comma-separated row data strings (null if using CSV files)
     * @param rowsFromCsv list of CSV files containing row data (null if using rows parameter)
     * @param rowIndex the starting row index for insertion
     * @param separator the delimiter used in row data (e.g., "," or ";")
     * @param eventHandler handler for sending progress notifications (null if no tracking needed)
     * @param taskId unique identifier for the async task (null if no tracking needed)
     * @param requestContext the request context for authentication and authorization
     * @throws RuntimeException if batch processing fails after all retry attempts
     * @throws RuntimeException if CSV files cannot be processed
     */
    public void addRowsInBatches(String driveId, String fileId, String tableName,
                                 List<String> rows, List<File> rowsFromCsv,
                                 int rowIndex, String separator,
                                 app.krista.extensions.util.EventHandler eventHandler, String taskId, RequestContext requestContext) {
        try {
            // Initialize batch processing context
            BatchProcessingContext context = initializeBatchProcessing(driveId, fileId, tableName, rows, rowsFromCsv, eventHandler, taskId, requestContext);

            // Process all batches
            processBatches(context, rowIndex, separator, requestContext);

            // Send success completion notification
            sendSuccessCompletionNotification(context);

        } catch (Exception e) {
            // Send failure completion notification if context is available
            long startTime = System.currentTimeMillis();
            if (eventHandler != null && taskId != null) {
                sendCompletionNotification(eventHandler, taskId, 0, 0, startTime, false, "Failed to add rows to table: " + e.getMessage() + ". Please check your data format and try again.");
            }
            throw e;
        }
    }

    /**
     * Initializes the batch processing context with all necessary data.
     */
    private BatchProcessingContext initializeBatchProcessing(String driveId, String fileId, String tableName,
                                                             List<String> rows, List<File> rowsFromCsv,
                                                             app.krista.extensions.util.EventHandler eventHandler, String taskId, RequestContext requestContext) {
        // Get all rows to process
        List<String> allRows = getRowsToProcess(rows, rowsFromCsv, eventHandler, taskId);

        // Get column count from table header
        int columnCount = getTableColumnCount(driveId, fileId, tableName, requestContext);

        // Create and return context
        BatchProcessingContext context = new BatchProcessingContext(driveId, fileId, tableName, allRows, columnCount, eventHandler, taskId);

        LOGGER.info("Processing {} rows in batches of {} with progress tracking", context.totalRows, context.batchSize);

        return context;
    }

    /**
     * Gets the rows to process from either the provided list or CSV files.
     */
    private List<String> getRowsToProcess(List<String> rows, List<File> rowsFromCsv,
                                          app.krista.extensions.util.EventHandler eventHandler, String taskId) {
        if (rows != null && !rows.isEmpty()) {
            return rows;
        }

        FileService fileService = new FileService(kristaMediaClient);
        try {
            return fileService.getRows(rowsFromCsv);
        } catch (IOException cause) {
            LOGGER.error("Failed to get rows from CSV files: {}", cause.getMessage(), cause);
            if (eventHandler != null && taskId != null) {
                sendCompletionNotification(eventHandler, taskId, 0, 0, System.currentTimeMillis(), false,
                        "Unable to process CSV files: " + cause.getMessage() + ". Please check that the files are valid CSV format and accessible.");
            }
            throw new RuntimeException("Unable to process CSV files. Please verify the files are in valid CSV format and accessible.", cause);
        }
    }

    /**
     * Gets the column count from the table header.
     */
    private int getTableColumnCount(String driveId, String fileId, String tableName, RequestContext requestContext) {
        try {
            WorkbookRange workbookRange = getWorkbook(driveId, fileId, requestContext).tables(tableName).headerRowRange().buildRequest().get();
            return workbookRange != null && workbookRange.columnCount != null ? workbookRange.columnCount : 0;
        } catch (Exception e) {
            LOGGER.error("Failed to get table column count: {}", e.getMessage(), e);
            throw new RuntimeException("Failed to get table column count", e);
        }
    }

    /**
     * Processes all batches with retry logic and progress tracking.
     */
    private void processBatches(BatchProcessingContext context, int rowIndex, String separator, RequestContext requestContext) {
        for (int i = 0; i < context.allRows.size(); i += context.batchSize) {
            int endIndex = Math.min(i + context.batchSize, context.allRows.size());
            List<String> batch = context.allRows.subList(i, endIndex);

            // Process single batch with retry logic
            processSingleBatch(context, batch, rowIndex + i, separator, i, endIndex, requestContext);

            // Update processed rows and check progress milestones
            context.processedRows += batch.size();
            checkAndSendProgressNotification(context);

            // Rate limiting delay
            addRateLimitingDelay(i, context);
        }

        LOGGER.info("Successfully processed {} rows in {} batches", context.totalRows,
                (context.totalRows + context.batchSize - 1) / context.batchSize);
    }

    /**
     * Processes a single batch with retry logic.
     */
    private void processSingleBatch(BatchProcessingContext context, List<String> batch, int batchRowIndex,
                                    String separator, int batchStartIndex, int endIndex, RequestContext requestContext) {
        final int MAX_RETRIES = 5;
        int retryCount = 0;
        boolean success = false;

        while (!success && retryCount <= MAX_RETRIES) {
            try {
                logBatchAttempt(retryCount, batchStartIndex, endIndex, context.totalRows, MAX_RETRIES);

                addRowsInTableBatch(context.driveId, context.fileId, context.tableName,
                        batch, context.columnCount, batchRowIndex, separator, requestContext);

                success = true;

            } catch (Exception e) {
                retryCount++;
                if (retryCount <= MAX_RETRIES) {
                    handleRetryableError(e, retryCount, batchStartIndex, endIndex);
                } else {
                    handleFinalBatchError(context, e, batchStartIndex, endIndex, MAX_RETRIES);
                }
            }
        }
    }

    /**
     * Checks progress milestones and sends notifications if needed.
     */
    private void checkAndSendProgressNotification(BatchProcessingContext context) {
        if (context.eventHandler != null && context.taskId != null &&
                context.currentMilestoneIndex < context.progressMilestones.length) {

            double currentProgress = (double) context.processedRows / context.totalRows * 100;
            if (currentProgress >= context.progressMilestones[context.currentMilestoneIndex]) {
                int milestone = context.progressMilestones[context.currentMilestoneIndex];
                String message = String.format("Processed %d of %d rows (%d%% complete)",
                        context.processedRows, context.totalRows, milestone);

                sendProgressNotification(context.eventHandler, context.taskId, milestone, context.processedRows,
                        context.totalRows, message, false);

                context.currentMilestoneIndex++;
                LOGGER.info("Progress milestone reached: {}% - {}", milestone, message);
            }
        }
    }

    /**
     * Adds rate limiting delay between batches.
     */
    private void addRateLimitingDelay(int currentIndex, BatchProcessingContext context) {
        if (currentIndex + context.batchSize < context.allRows.size()) {
            try {
                Thread.sleep(2000);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new RuntimeException("Thread interrupted during rate limiting delay", e);
            }
        }
    }

    /**
     * Logs batch processing attempt.
     */
    private void logBatchAttempt(int retryCount, int startIndex, int endIndex, int totalRows, int maxRetries) {
        if (retryCount > 0) {
            LOGGER.info("Retrying batch {}-{} of {} (attempt {} of {})",
                    startIndex + 1, endIndex, totalRows, retryCount + 1, maxRetries + 1);
        } else {
            LOGGER.info("Processing batch {}-{} of {}", startIndex + 1, endIndex, totalRows);
        }
    }

    /**
     * Handles retryable errors with appropriate delays.
     */
    private void handleRetryableError(Exception e, int retryCount, int startIndex, int endIndex) {
        long retryDelay = getRetryDelay(e, retryCount);
        LOGGER.warn("Failed to process batch {}-{} on attempt {} ({}): {}. Retrying in {}ms...",
                startIndex + 1, endIndex, retryCount, e.getClass().getSimpleName(), e.getMessage(), retryDelay);
        try {
            Thread.sleep(retryDelay);
        } catch (InterruptedException ie) {
            Thread.currentThread().interrupt();
            throw new RuntimeException("Thread interrupted during retry delay", ie);
        }
    }

    /**
     * Handles final batch error after all retries exhausted.
     */
    private void handleFinalBatchError(BatchProcessingContext context, Exception e, int startIndex, int endIndex, int maxRetries) {
        LOGGER.error("Failed to process batch {}-{} after {} attempts ({}): {}",
                startIndex + 1, endIndex, maxRetries + 1, e.getClass().getSimpleName(), e.getMessage());

        // Send failure notification
        if (context.eventHandler != null && context.taskId != null) {
            sendCompletionNotification(context.eventHandler, context.taskId, context.totalRows, context.processedRows,
                    context.startTime, false, "Batch processing failed: " + e.getMessage());
        }

        throw new RuntimeException("Batch processing failed at rows " + (startIndex + 1) + "-" + endIndex +
                " after " + (maxRetries + 1) + " attempts", e);
    }

    /**
     * Sends success completion notification.
     */
    private void sendSuccessCompletionNotification(BatchProcessingContext context) {
        if (context.eventHandler != null && context.taskId != null) {
            // Send 100% progress notification first
            String progressMessage = String.format("Processed %d of %d rows (100%% complete)",
                    context.processedRows, context.totalRows);
            sendProgressNotification(context.eventHandler, context.taskId, 100, context.processedRows,
                    context.totalRows, progressMessage, true);

            try {
                Thread.sleep(5000);
            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
                throw new RuntimeException("Thread interrupted during retry delay", ie);
            }

            // Then send completion notification to result available request
            sendCompletionNotification(context.eventHandler, context.taskId, context.totalRows, context.processedRows,
                    context.startTime, true, "Batch processing completed successfully");
        }
    }

    /**
     * Sends progress notification through the event handler.
     */
    private void sendProgressNotification(app.krista.extensions.util.EventHandler eventHandler,
                                          String taskId, int progressPercentage, int currentIndex,
                                          int totalRows, String message, boolean completed) {
        try {
            app.krista.model.base.FreeForm progressData = new app.krista.model.base.FreeForm();
            progressData.put("Message", "Paragraph", message);
            progressData.put("Current Index", "Number", currentIndex);
            progressData.put("Progress", "Percentage", progressPercentage / 100.0); // Convert to decimal for percentage field
            progressData.put("Completed", "Boolean", completed);
            progressData.put("Total Rows", "Number", totalRows);

            app.krista.model.base.FreeForm eventData = new app.krista.model.base.FreeForm();
            eventData.put("Data", "[ Composite ]", progressData);

            eventHandler.handleEvent(taskId, eventData);
            LOGGER.info("Progress notification sent: {}% - {}", progressPercentage, message);
        } catch (Exception e) {
            LOGGER.warn("Failed to send progress notification: {}", e.getMessage());
        }
    }

    /**
     * Sends completion notification to addRowsAsyncToTableResultAvailable through the event handler.
     */
    private void sendCompletionNotification(app.krista.extensions.util.EventHandler eventHandler,
                                            String taskId, int totalRows, int rowsAdded,
                                            long startTime,
                                            boolean success, String message) {
        try {
            long endTime = System.currentTimeMillis();
            long duration = endTime - startTime;

            app.krista.model.base.FreeForm completionData = new app.krista.model.base.FreeForm();
            completionData.put("Total Rows", "Number", totalRows);
            completionData.put("Rows Added", "Number", rowsAdded);
            completionData.put("Success", "Boolean", success);
            completionData.put("Message", "Paragraph", message);
            completionData.put("Start Time", "Time", new java.util.Date(startTime));
            completionData.put("End Time", "Time", new java.util.Date(endTime));
            completionData.put("Duration", "Number", duration);
            completionData.put("Duration Text", "Text", formatDuration(duration));

            app.krista.model.base.FreeForm eventData = new app.krista.model.base.FreeForm();
            eventData.put("Data", "[ Composite ]", completionData);

            // Use the same task ID but mark the data to indicate it's a completion event
            completionData.put("EventType", "Text", "COMPLETION");
            eventHandler.handleEvent(taskId + "_completion", eventData);
            LOGGER.info("Completion notification sent: {} rows processed successfully in {}ms", rowsAdded, duration);
        } catch (Exception e) {
            LOGGER.warn("Failed to send completion notification: {}", e.getMessage());
        }
    }

    /**
     * Formats duration in milliseconds to human-readable text.
     */
    private String formatDuration(long durationMs) {
        if (durationMs < 1000) {
            return durationMs + "ms";
        } else if (durationMs < 60000) {
            return String.format("%.1fs", durationMs / 1000.0);
        } else {
            long minutes = durationMs / 60000;
            long seconds = (durationMs % 60000) / 1000;
            return String.format("%dm %ds", minutes, seconds);
        }
    }

    /**
     * Determines the retry delay based on the exception type and retry attempt number.
     * Network-related exceptions get longer delays.
     */
    private long getRetryDelay(Exception e, int retryCount) {
        // Check if it's a network/connectivity related exception
        if (isNetworkRelatedError(e)) {
            // For network errors, use progressive longer delays: 10s, 20s, 30s, 45s, 60s
            return switch (retryCount) {
                case 1 -> 5000L;   // 5 seconds
                case 2 -> 10000L;  // 10 seconds
                case 3 -> 15000L;  // 15 seconds
                case 4 -> 20000L;  // 20 seconds
                case 5 -> 25000L;  // 25 seconds
                default -> 60000L; // Cap at 60 seconds for any additional retries
            };
        } else {
            // For other errors, use progressive delays: 5s, 10s, 15s, 20s, 25s
            return switch (retryCount) {
                case 1 -> 5000L;   // 5 seconds
                case 2 -> 10000L;  // 10 seconds
                case 3 -> 15000L;  // 15 seconds
                case 4 -> 20000L;  // 20 seconds
                case 5 -> 25000L;  // 25 seconds
                default -> 25000L; // Cap at 25 seconds for any additional retries
            };
        }
    }

    /**
     * Checks if the exception is related to network connectivity issues.
     */
    private boolean isNetworkRelatedError(Exception e) {
        if (e == null) return false;

        String message = e.getMessage();
        String className = e.getClass().getSimpleName();

        // Check for common network-related error indicators
        return className.contains("ClientException") ||
                className.contains("ConnectException") ||
                className.contains("SocketTimeoutException") ||
                className.contains("UnknownHostException") ||
                (message != null && (
                        message.toLowerCase().contains("timeout") ||
                                message.toLowerCase().contains("connection") ||
                                message.toLowerCase().contains("network") ||
                                message.toLowerCase().contains("unreachable") ||
                                message.toLowerCase().contains("error executing the request")
                ));
    }

    /**
     * Adds a batch of rows to an Excel table using the Microsoft Graph API.
     * This is an internal method used by the batch processing logic.
     *
     * @param driveId the ID of the drive containing the workbook
     * @param fileID the ID of the Excel file
     * @param tableName the name of the table to add rows to
     * @param rows list of comma-separated row data strings
     * @param columnCount the number of columns in the table
     * @param index the starting row index for insertion
     * @param separator the delimiter used in row data
     * @param requestContext the request context for authentication
     * @throws RuntimeException if the batch operation fails
     */
    private void addRowsInTableBatch(String driveId, String fileID, String tableName,
                                     List<String> rows, int columnCount, Integer index, String separator, RequestContext requestContext) {
        try {

            JsonElement inputRows = Tables.transformRows(rows, columnCount, separator);

            WorkbookTableRowAddParameterSet params = WorkbookTableRowAddParameterSet
                    .newBuilder()
                    .withIndex(index)
                    .withValues(inputRows)
                    .build();

            getWorkbook(driveId, fileID, requestContext).tables(tableName).rows().add(params).buildRequest().post();
        } catch (NullPointerException | GraphServiceException cause) {
            LOGGER.error("Failed to add rows in table batch: {}", cause.getMessage(), cause);
            throw new RuntimeException("Failed to add rows in table batch", cause);
        }
    }

    /**
     * Adds rows to an Excel table synchronously and returns the updated table data.
     * Accepts row data either as a list of strings or from CSV files.
     *
     * @param driveId the ID of the drive containing the workbook
     * @param fileID the ID of the Excel file
     * @param tableName the name of the table to add rows to
     * @param rows list of comma-separated row data strings (null if using CSV files)
     * @param rowsFromCsvFile list of CSV files containing row data (null if using rows parameter)
     * @param index the starting row index for insertion (null to append at end)
     * @param separator the delimiter used in row data (e.g., "," or ";")
     * @return list of FreeForm objects representing the updated table data, empty list on error
     */
    public List<FreeForm> addRowsInTable(String driveId, String fileID, String tableName,
                                         List<String> rows, List<File> rowsFromCsvFile, Integer index, String separator) {
        try {
            List<String> rowsToAdd = rows;
            if (rows == null || rows.isEmpty()) {
                FileService fileService = new FileService(kristaMediaClient);
                rowsToAdd = fileService.getRows(rowsFromCsvFile);
            }

            WorkbookRange workbookRange = getWorkbook(driveId, fileID).tables(tableName).headerRowRange().buildRequest().get();
            assert workbookRange != null;
            Integer columnCount = workbookRange.columnCount;

            JsonElement inputRows = Tables.transformRows(rowsToAdd, columnCount, separator);

            WorkbookTableRowAddParameterSet params = WorkbookTableRowAddParameterSet
                    .newBuilder()
                    .withIndex(index)
                    .withValues(inputRows)
                    .build();

            getWorkbook(driveId, fileID).tables(tableName).rows().add(params).buildRequest().post();
            return tableService.getTableData(driveId, fileID, tableName);
        } catch (NullPointerException | GraphServiceException | IOException cause) {
            LOGGER.error("Failed to add rows in table: {}", cause.getMessage(), cause);
            return Collections.emptyList();
        }
    }

    @NotNull
    private WorkbookRequestBuilder getWorkbook(String driveId, String fileID) {
        return getClientProvider().getDriveRequestBuilder(driveId).items(fileID).workbook();
    }

    @NotNull
    private WorkbookRequestBuilder getWorkbook(String driveId, String fileID, RequestContext requestContext) {
        return getClientProvider().getDriveRequestBuilder(driveId, requestContext).items(fileID).workbook();
    }

    /**
     * Calculates the sum of specified columns in an Excel table.
     * Supports multiple columns and comma-separated column names within a single string.
     *
     * @param driveId the ID of the drive containing the workbook
     * @param fileID the ID of the Excel file
     * @param tableName the name of the table containing the columns
     * @param columns list of column names to sum (supports comma-separated values)
     * @return the sum of all specified columns as a Double, or 0.0 on error
     */
    public Double tableColumnsSum(String driveId, String fileID,
                                  String tableName, List<String> columns) {
        try {
            WorkbookRequestBuilder workbookRequestBuilder = getWorkbook(driveId, fileID);
            List<Integer> numbers = new ArrayList<>();
            for (String column : columns) {
                // Split comma-separated columns within a single string
                String[] splitColumns = column.split(",");
                for (String singleColumn : splitColumns) {
                    String trimmedColumn = singleColumn.trim();
                    if (!trimmedColumn.isEmpty()) {
                        WorkbookTableColumnDataBodyRangeRequestBuilder rangeReqBuilder = workbookRequestBuilder.tables(tableName).columns(trimmedColumn)
                                .dataBodyRange();
                        WorkbookRange wbRange = rangeReqBuilder.buildRequest().get();
                        if (wbRange == null || wbRange.values == null) {
                            LOGGER.error("Workbook range is null");
                            return null;
                        }
                        numbers.addAll(DataParser.convertRangeIntoFlatList(wbRange.values));
                    }
                }
            }

            JsonElement numbersJsonElement = GSON.toJsonTree(numbers);
            return Worksheets.getSum(numbersJsonElement);
        } catch (GraphServiceException cause) {
            LOGGER.error("Cause:{}", cause.getMessage(), cause);
            return null;
        }
    }

    public FreeForm executeFormula(String driveId, String fileId, String formula,
                                   String formulaName, String comment) {
        String sanitizedName = sanitizeFormulaName(formulaName);

        if (!isValidInput(fileId, formula, formulaName, sanitizedName)) {
            return new FreeForm();
        }

        try {
            WorkbookNamedItemAddFormulaLocalParameterSet paramSet = buildFormulaParameterSet(sanitizedName, formula, comment);
            WorkbookNamedItem formulaItem = postFormula(driveId, fileId, paramSet);
            if (formulaItem == null || formulaItem.value == null) {
                LOGGER.error("Formula item is null");
                return new FreeForm();
            }
            LOGGER.info("Formula '{}' executed successfully.", sanitizedName);
            return DataParser.convertToFreeForm(formula, formulaItem.value);
        } catch (GraphServiceException cause) {
            return handleGraphServiceException(formula, cause);
        } catch (Exception cause) {
            throw new IllegalArgumentException(cause);
        } finally {
            cleanupFormulaName(driveId, fileId, sanitizedName);
        }
    }

    private String sanitizeFormulaName(String formulaName) {
        return formulaName != null ? formulaName.replaceAll("[^a-zA-Z0-9_]", "_") : null;
    }

    private boolean isValidInput(String fileId, String formula, String formulaName, String sanitizedName) {
        boolean valid = fileId != null && !fileId.trim().isEmpty() &&
                formula != null && !formula.trim().isEmpty() &&
                formulaName != null && !formulaName.trim().isEmpty() &&
                sanitizedName != null && !sanitizedName.isEmpty();

        if (!valid) {
            LOGGER.error("Invalid parameters: fileId={}, formula={}, formulaName={}", fileId, formula, formulaName);
        }

        return valid;
    }

    private WorkbookNamedItemAddFormulaLocalParameterSet buildFormulaParameterSet(
            String sanitizedName, String formula, String comment) {

        var builder = WorkbookNamedItemAddFormulaLocalParameterSet.newBuilder()
                .withName(sanitizedName)
                .withFormula(formula);

        if (comment != null && !comment.trim().isEmpty()) {
            builder = builder.withComment(comment);
        }

        return builder.build();
    }

    private WorkbookNamedItem postFormula(String driveId, String fileId,
                                          WorkbookNamedItemAddFormulaLocalParameterSet paramSet) {
        return getWorkbook(driveId, fileId)
                .names()
                .addFormulaLocal(paramSet)
                .buildRequest()
                .post();
    }

    private FreeForm handleGraphServiceException(String formula, GraphServiceException e) {
        LOGGER.error("Failed to execute formula: {} - Code: {}, Message: {}",
                formula, e.getServiceError() != null ? e.getServiceError().code : "Unknown",
                e.getServiceError() != null ? e.getServiceError().message : e.getMessage(), e);

        FreeForm errorResult = new FreeForm();
        String errorMessage = e.getServiceError() != null ? e.getServiceError().message : e.getMessage();
        errorResult.put("Error", "Text", "Unable to execute Excel formula: " + errorMessage + ". " +
                "Please check that your formula syntax is correct, all referenced cells exist, and you have permission to modify the workbook.");
        return errorResult;
    }

    private FreeForm handleGenericException(Exception e) {
        LOGGER.error("Unexpected error executing formula: {}", e.getMessage(), e);
        FreeForm errorResult = new FreeForm();
        errorResult.put("Error", "Text", "An unexpected error occurred while executing the formula: " + e.getMessage() + ". " +
                "Please check your formula syntax and try again. If the issue continues, contact your system administrator.");
        return errorResult;
    }

    private void cleanupFormulaName(String driveId, String fileId, String sanitizedName) {
        if (sanitizedName == null || sanitizedName.isEmpty()) return;

        try {
            getWorkbook(driveId, fileId).names(sanitizedName)
                    .buildRequest()
                    .delete();
            LOGGER.debug("Cleaned up formula name: {}", sanitizedName);
        } catch (Exception e) {
            LOGGER.warn("Failed to clean up formula name '{}': {}", sanitizedName, e.getMessage());
        }
    }

    public FreeForm addFormula(String workbookID, String sheetName, String cellID, String formula) {
        try {
            WorkbookWorksheetRangeParameterSet rangeParameterSet = Workbooks.getWorkbookWorksheetRangeParameterSet(cellID);

            WorkbookWorksheetRangeRequestBuilder rangeRequestBuilder = getWorkbook(null, workbookID)
                    .worksheets(sheetName)
                    .range(rangeParameterSet);

            if (rangeRequestBuilder.getRequestUrl() == null || rangeRequestBuilder.getClient() == null) {
                LOGGER.error("Range request is null");
                return new FreeForm();
            }

            WorkbookWorksheetRangeRequestBuilderV2 rangeReqBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    rangeRequestBuilder.getRequestUrl(), rangeRequestBuilder.getClient(),
                    rangeRequestBuilder.getOptions(), rangeParameterSet);

            WorkbookRange workbookRange = new WorkbookRange();
            workbookRange.address = sheetName + Constants.SHEET_RANGE_SEPARATOR + cellID;
            workbookRange.formulas = GSON.toJsonTree(formula);
            rangeReqBuilderV2.buildRequest().patch(workbookRange);

            FreeForm cellInformation = new FreeForm();
            cellInformation.put(Constants.CELL_ID, Constants.TEXT, cellID);
            cellInformation.put(Constants.FORMULA, Constants.TEXT, formula);
            return cellInformation;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to add formula: {}", cause.getMessage(), cause);
            return new FreeForm();
        }
    }

    public Double worksheetRangeSum(String driveId, String workbookId,
                                    String worksheetName, String range) {
        try {
            WorkbookRequestBuilder workbookRequestBuilder = getWorkbook(driveId, workbookId);
            WorkbookWorksheetRangeParameterSet rangeParameterSet = Workbooks.getWorkbookWorksheetRangeParameterSet(range);

            WorkbookWorksheetRangeRequestBuilder rangeReqBuilder = workbookRequestBuilder
                    .worksheets(worksheetName)
                    .range(rangeParameterSet);
            if (rangeReqBuilder.getRequestUrl() == null || rangeReqBuilder.getClient() == null) {
                LOGGER.error("Range request is null");
                return null;
            }
            WorkbookWorksheetRangeRequestBuilderV2 rangeRequestBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    rangeReqBuilder.getRequestUrl(), rangeReqBuilder.getClient(),
                    rangeReqBuilder.getOptions(), rangeParameterSet);
            WorkbookRange workbookRange = rangeRequestBuilderV2.buildRequest().get();
            if (workbookRange == null || workbookRange.values == null) {
                LOGGER.error("Workbook range is null");
                return null;
            }
            JsonElement numbers = DataParser.convertRangeIntoFlatRange(workbookRange.values);
            return Worksheets.getSum(numbers);

        } catch (GraphServiceException cause) {
            LOGGER.error("Cause: {}", cause.getMessage(), cause);
            return null;
        }
    }

    public FreeForm vLookup(String driveId, String workbookID, String lookupValue,
                            String sheetName, String range, Double columnIndex, Boolean exactMatch) {
        try {
            JsonElement lookupVal = GSON.toJsonTree(lookupValue);
            JsonElement colIndex = GSON.toJsonTree(columnIndex);
            JsonElement rangeLookup = GSON.toJsonTree(exactMatch);

            WorkbookRequestBuilder workbook = getWorkbook(driveId, workbookID);
            WorkbookWorksheetRangeParameterSet parameterSet = Workbooks.getWorkbookWorksheetRangeParameterSet(range.trim());
            WorkbookWorksheetRangeRequestBuilder rangeBuilder = workbook.worksheets(sheetName.trim()).range(parameterSet);

            WorkbookWorksheetRangeRequestBuilderV2 rangeRequestBuilderV2 = new WorkbookWorksheetRangeRequestBuilderV2(
                    Objects.requireNonNull(rangeBuilder.getRequestUrl()), Objects.requireNonNull(rangeBuilder.getClient()), rangeBuilder.getOptions(), parameterSet);
            JsonElement rangeJsonElem = Objects.requireNonNull(rangeRequestBuilderV2.buildRequest().get()).values;
            WorkbookFunctionsVlookupParameterSet vLookupParameterSet = Workbooks.getvLookupParameterSet(lookupVal, colIndex, rangeLookup, rangeJsonElem);

            WorkbookFunctionResult functionResult = workbook.functions().vlookup(vLookupParameterSet).buildRequest().post();
            if (functionResult == null || functionResult.value == null) {
                LOGGER.error("Function result is null");
                return new FreeForm();
            }
            Object value = DataParser.getCellData(functionResult.value);

            return Workbooks.getFreeForm(lookupValue, value);
        } catch (GraphServiceException cause) {
            LOGGER.error("Cause: {}", cause.getMessage(), cause);
            return new FreeForm();
        }
    }

    /**
     * Checks if a drive is accessible with the given drive ID.
     *
     * @param driveId The ID of the drive to check
     * @return true if the drive is accessible, false otherwise
     */
    public boolean isDriveAccessible(String driveId) {
        if (driveId == null || driveId.trim().isEmpty()) {
            return false;
        }

        try {
            getClientProvider().getDriveRequestBuilder(driveId).buildRequest().get();
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error accessing drive with ID {}: {}", driveId, cause.getMessage());
            return false;
        }
    }

    /**
     * Checks if a file is accessible with the given drive ID and file ID.
     *
     * @param driveId The ID of the drive containing the file
     * @param fileId  The ID of the file to check
     * @return true if the file is accessible, false otherwise
     */
    public boolean isFileAccessible(String driveId, String fileId) {
        try {
            String trimmedDriveId = (driveId != null && !driveId.trim().isEmpty()) ? driveId : null;
            getClientProvider().getDriveRequestBuilder(trimmedDriveId).items(fileId).buildRequest().get();
            return true;
        } catch (GraphServiceException cause) {
            LOGGER.error("Error accessing file with ID {} in drive {}: {}", fileId, driveId, cause.getMessage());
            return false;
        }
    }

    /**
     * Checks if a table is accessible with the given drive ID, file ID, and table name.
     *
     * @param driveId   The ID of the drive containing the file
     * @param fileId    The ID of the file containing the table
     * @param tableName The name of the table to check
     * @return true if the table is accessible, false otherwise
     */
    public boolean isTableAccessible(String driveId, String fileId, String tableName) {
        if (fileId == null || fileId.trim().isEmpty() ||
                tableName == null || tableName.trim().isEmpty()) {
            return false;
        }

        try {
            String trimmedDriveId = (driveId != null && !driveId.trim().isEmpty()) ? driveId : null;
            getClientProvider().getDriveRequestBuilder(trimmedDriveId).items(fileId).workbook()
                    .tables(tableName).buildRequest().get();
            return true;
        } catch (GraphServiceException cause) {
            LOGGER.error("Error accessing table '{}' in file {} in drive {}: {}",
                    tableName, fileId, driveId, cause.getMessage());
            return false;
        }
    }

    /**
     * Checks if a column exists in a table with the given drive ID, file ID, table name, and column name.
     *
     * @param driveId    The ID of the drive containing the file
     * @param fileId     The ID of the file containing the table
     * @param tableName  The name of the table containing the column
     * @param columnName The name of the column to check
     * @return true if the column exists in the table, false otherwise
     */
    public boolean isColumnExistsInTable(String driveId, String fileId, String tableName, String columnName) {
        if (fileId == null || fileId.trim().isEmpty() || tableName == null || tableName.trim().isEmpty() ||
                columnName == null || columnName.trim().isEmpty()) {
            return false;
        }

        try {
            String trimmedDriveId = (driveId != null && !driveId.trim().isEmpty()) ? driveId : null;
            getClientProvider().getDriveRequestBuilder(trimmedDriveId).items(fileId).workbook().tables(tableName).columns(columnName).buildRequest().get();
            return true;
        } catch (GraphServiceException cause) {
            if (cause.getResponseCode() == 404) {
                LOGGER.debug("Column '{}' does not exist in table '{}' in file {} in drive {}",
                        columnName, tableName, fileId, driveId);
            } else {
                LOGGER.error("Error checking if column exists: {}", cause.getMessage(), cause);
            }
            return false;
        } catch (Exception e) {
            LOGGER.error("Unexpected error checking if column exists: {}", e.getMessage(), e);
            return false;
        }
    }
}
