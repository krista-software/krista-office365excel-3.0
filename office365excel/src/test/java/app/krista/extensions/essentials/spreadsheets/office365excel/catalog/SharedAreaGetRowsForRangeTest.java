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
import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.SharedOperationsFacade;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.util.EventHandler;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.ksdk.entities.Entities;
import app.krista.model.base.FreeForm;
import com.kristasoft.common.holders.ThreadLocalProxy;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.ArgumentCaptor;
import org.mockito.Mock;
import org.mockito.MockedStatic;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.*;

/**
 * Unit tests for SharedArea getRowsForRange related methods.
 * Ensures 100% code coverage.
 */
@ExtendWith(MockitoExtension.class)
@DisplayName("SharedArea - Get Rows For Range Tests")
public class SharedAreaGetRowsForRangeTest {

    private static final String DRIVE_ID = "testDriveId";
    private static final String FILE_ID = "testFileId";
    private static final String SHEET_NAME = "testSheetName";
    private static final String RANGE = "A1:B10";

    @Mock private DriveService driveService;
    @Mock private RestClientProvider restClientProvider;
    @Mock private RequestContext requestContext;
    @Mock private EventHandler eventHandler;
    @Mock private AuthorizationContext authorizationContext;
    @Mock private Entities entities;
    @Mock private WorkbookService workbookService;
    @Mock private TableService tableService;
    @Mock private WorksheetService worksheetService;
    @Mock private WorksheetOperation worksheetOperation;
    @Mock private StringUtil stringUtil;
    @Mock private SheetFormattingService sheetFormattingService;
    @Mock private SharedOperationsFacade sharedOperationsFacade;
    @Mock private Office365ExcelAuditService auditService;
    @Mock private GraphServiceClientProviderFactory providerFactory;
    @Mock private Invoker invoker;

    private SharedArea sharedArea;

    @BeforeEach
    public void setUp() {
        when(invoker.getInvokerId()).thenReturn("test-invoker");
        sharedArea = new SharedArea(
                driveService, restClientProvider, requestContext,
                eventHandler, authorizationContext, entities,
                workbookService, tableService, worksheetService,
                worksheetOperation, stringUtil, sheetFormattingService,
                sharedOperationsFacade, auditService, providerFactory, invoker
        );

        sharedArea.executorService = mock(java.util.concurrent.ExecutorService.class);
        lenient().doAnswer(invocation -> {
            try {
                ((Runnable) invocation.getArgument(0)).run();
            } catch (Exception e) {
                // Background executors swallow exceptions on submit() unless future.get() is called
            }
            return java.util.concurrent.CompletableFuture.completedFuture(null);
        }).when(sharedArea.executorService).submit(any(Runnable.class));
    }

    // --- getRowsForRange ---

    @Test
    @DisplayName("getRowsForRange - passes all parameters including allowRetry (null)")
    public void testGetRowsForRange_NullAllowRetry() {
        ExtensionResponse mockResponse = mock(ExtensionResponse.class);
        when(sharedOperationsFacade.getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null))
                .thenReturn(mockResponse);

        ExtensionResponse response = sharedArea.getRowsForRange(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

        assertEquals(mockResponse, response);
        verify(sharedOperationsFacade).getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);
    }

    @Test
    @DisplayName("getRowsForRange - passes all parameters including allowRetry (true)")
    public void testGetRowsForRange_TrueAllowRetry() {
        ExtensionResponse mockResponse = mock(ExtensionResponse.class);
        when(sharedOperationsFacade.getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true))
                .thenReturn(mockResponse);

        ExtensionResponse response = sharedArea.getRowsForRange(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

        assertEquals(mockResponse, response);
        verify(sharedOperationsFacade).getRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);
    }

    // --- getRowsForRangeAsync ---

    @Test
    @DisplayName("getRowsForRangeAsync - Success path triggers eventHandler asynchronously")
    public void testGetRowsForRangeAsync_Success() {
        // Arrange Mock Contexts
        Map<Class<?>, Object> dummyThreadLocals = Collections.emptyMap();

        try (MockedStatic<ThreadLocalProxy> proxyMock = mockStatic(ThreadLocalProxy.class)) {
            proxyMock.when(ThreadLocalProxy::getAll).thenReturn(dummyThreadLocals);

            FreeForm expectedRow = mock(FreeForm.class);
            List<FreeForm> expectedRows = Collections.singletonList(expectedRow);

            when(worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE)).thenReturn(expectedRows);

            // Act
            String taskId = sharedArea.getRowsForRangeAsync(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            // Assert
            assertNotNull(taskId);

            // Verify
            ArgumentCaptor<FreeForm> freeFormCaptor = ArgumentCaptor.forClass(FreeForm.class);
            verify(eventHandler).handleEvent(eq(taskId), freeFormCaptor.capture());

            FreeForm eventData = freeFormCaptor.getValue();
            assertNotNull(eventData);
            assertEquals(expectedRows, eventData.get("Rows"));

            // Verify ThreadLocal proxy static calls were made
            proxyMock.verify(() -> ThreadLocalProxy.setAll(dummyThreadLocals), times(1));
            proxyMock.verify(ThreadLocalProxy::removeAll, times(1));
        }
    }

    @Test
    @DisplayName("getRowsForRangeAsync - Exception path throws RuntimeException inside thread layer")
    public void testGetRowsForRangeAsync_Exception() {
        // Arrange
        try (MockedStatic<ThreadLocalProxy> proxyMock = mockStatic(ThreadLocalProxy.class)) {
            proxyMock.when(ThreadLocalProxy::getAll).thenReturn(Collections.emptyMap());

            when(worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE))
                    .thenThrow(new IllegalArgumentException("Mock graph exception"));

            // Act
            String taskId = sharedArea.getRowsForRangeAsync(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);

            // Assert
            assertNotNull(taskId);

            // Allow thread to process and fail (eventHandler should NEVER be called on exception)
            verify(eventHandler, never()).handleEvent(any(), any());

            // Verify removeAll was still called inside finally block
            proxyMock.verify(ThreadLocalProxy::removeAll, times(1));
        }
    }

    // --- getRowsForRangeResult ---

    @Test
    @DisplayName("getRowsForRangeResult - returns rows when eventName equals taskID")
    public void testGetRowsForRangeResult_ValidTaskId() {
        // Arrange
        String taskId = UUID.randomUUID().toString();
        String eventName = taskId; // Matching eventName

        FreeForm mockRow = mock(FreeForm.class);
        List<FreeForm> expectedRows = Collections.singletonList(mockRow);

        FreeForm eventData = mock(FreeForm.class);
        when(eventData.get("Rows")).thenReturn(expectedRows);

        // Act
        List<FreeForm> result = sharedArea.getRowsForRangeResult(eventName, eventData, taskId);

        // Assert
        assertEquals(expectedRows, result);
    }

    @Test
    @DisplayName("getRowsForRangeResult - throws IllegalStateException when eventName does not match taskID")
    public void testGetRowsForRangeResult_InvalidTaskId() {
        // Arrange
        String taskId = UUID.randomUUID().toString();
        String eventName = UUID.randomUUID().toString(); // Non-matching eventName

        FreeForm eventData = mock(FreeForm.class);

        // Act & Assert
        IllegalStateException exception = assertThrows(IllegalStateException.class,
                () -> sharedArea.getRowsForRangeResult(eventName, eventData, taskId));

        assertEquals("Invalid task ID", exception.getMessage());
    }
}
