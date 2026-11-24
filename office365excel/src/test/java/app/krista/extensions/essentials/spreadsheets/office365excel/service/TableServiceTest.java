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
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonElement;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.requests.*;
import junit.framework.TestCase;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.io.IOException;
import java.util.List;

import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

public class TableServiceTest extends TestCase {
    @Mock
    private TableService tableService;

    @Mock
    private DriveRequestBuilder driveRequestBuilder;
    @Mock
    private DriveItemRequestBuilder driveItemRequestBuilder;
    @Mock
    private WorkbookRequestBuilder workbookRequestBuilder;
    @Mock
    private WorkbookTableRequestBuilder workbookTableRequestBuilder;

    @Mock
    private WorkbookTableRangeRequest workbookTableRangeRequest;
    @Mock
    private GraphServiceClientProvider provider;
    @Mock
    private GraphServiceClient client;
    @Mock
    private UserRequestBuilder userRequestBuilder;

    public void setUp() {
        MockitoAnnotations.openMocks(this);
    }

    @Test
    public void testGetTableData() throws IOException {

        WorkbookTableRangeRequestBuilder tableRangeRequestBuilder = mock(WorkbookTableRangeRequestBuilder.class);
        WorkbookRange tableRange = new WorkbookRange();

        String workbookID = "workbookID";
        String tableName = "tableName";

        String jsonData = "[\n" +
                "        [\n" +
                "            \"My Date\",\n" +
                "            \"Origin Tare Weight\",\n" +
                "            \"Origin Gross Weight\",\n" +
                "            \"Origin Net Weight\",\n" +
                "            \"Transporter\"\n" +
                "        ],\n" +
                "        [\n" +
                "            \"4/2/2022\",\n" +
                "            \"12424\",\n" +
                "            \"22.00%\",\n" +
                "            \"29.74\",\n" +
                "            \"FORTUNE LOGISTICS\"\n" +
                "        ],\n" +
                "        [\n" +
                "            \"4/3/2022\",\n" +
                "            \"12.86\",\n" +
                "            \"41.99%\",\n" +
                "            \"29.13\",\n" +
                "            \"FORTUNE LOGISTICS\"\n" +
                "        ]\n" +
                "]";

        tableRange.text = Constants.GSON.fromJson(jsonData, JsonElement.class);

        when(provider.getGraphServiceClient()).thenReturn(client);
        when(client.me()).thenReturn(userRequestBuilder);
        when(userRequestBuilder.drives(anyString())).thenReturn(driveRequestBuilder);

        when(driveRequestBuilder.items(workbookID)).thenReturn(driveItemRequestBuilder);
        when(driveItemRequestBuilder.workbook()).thenReturn(workbookRequestBuilder);
        when(workbookRequestBuilder.tables(tableName)).thenReturn(workbookTableRequestBuilder);
        when(workbookTableRequestBuilder.range()).thenReturn(tableRangeRequestBuilder);

        when(tableRangeRequestBuilder.buildRequest()).thenReturn(workbookTableRangeRequest);
        when(workbookTableRangeRequest.get()).thenReturn(tableRange);

        final List<FreeForm> output = tableService.getTableData("driveId", workbookID, tableName);
        assertNotNull(output);

    }
}

