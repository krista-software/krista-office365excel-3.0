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
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.ItemReference;
import com.microsoft.graph.requests.*;
import junit.framework.TestCase;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.io.IOException;

import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.when;

public class WorkbookServiceTest extends TestCase {

    @Mock
    private GraphServiceClient graphServiceClient;

    @Mock
    private UserRequestBuilder userRequestBuilder;

    @Mock
    private DriveRequestBuilder drive;
    @Mock
    private DriveItemRequestBuilder root;

    @Mock
    private DriveItemRequestBuilder driveItemRequestBuilder;

    @Mock
    private DriveItemRequest driveItemRequest;

    @Mock
    private GraphServiceClientProvider clientProvider;
    @Mock
    private WorkbookOperation workbookOperation;

    private WorkbookService workbookService;
    private KristaMediaClient kristaMediaClient;
    private TableService tableService;

    public void setUp() {
        MockitoAnnotations.openMocks(this);
        workbookService = new WorkbookService(workbookOperation, clientProvider, kristaMediaClient, tableService);
    }

    @Test
    public void testGetWorkbookDetails() throws IOException {

        // Given
        String workbookFolderPath = "excelTest";
        String workbookName = "excelFile";

        // Mocked Drive Item
        DriveItem outputFile = new DriveItem();

        outputFile.id = "excelFileId";
        ItemReference itemReference = new ItemReference();
        itemReference.driveId = "oneDriveId";
        outputFile.parentReference = itemReference;
        outputFile.name = "excelFile";
        outputFile.webUrl = "webUrl";

        // Stub

        when(clientProvider.getGraphServiceClient()).thenReturn(graphServiceClient);
        when(clientProvider.getDriveRequestBuilder(null)).thenReturn(drive);
        when(graphServiceClient.me()).thenReturn(userRequestBuilder);
        when(userRequestBuilder.drive()).thenReturn(drive);
        when(drive.root()).thenReturn(root);
        when(root.itemWithPath(anyString())).thenReturn(driveItemRequestBuilder);
        when(driveItemRequestBuilder.buildRequest()).thenReturn(driveItemRequest);
        when(driveItemRequest.get()).thenReturn(outputFile);

        // When
        final ExcelFile workbookDetails = workbookService.getWorkbookDetails(workbookFolderPath, workbookName);

        assertNotNull(workbookDetails);
        assertEquals("excelFile", workbookDetails.fileName);
        assertEquals("excelFileId", workbookDetails.fileId);
    }
}
