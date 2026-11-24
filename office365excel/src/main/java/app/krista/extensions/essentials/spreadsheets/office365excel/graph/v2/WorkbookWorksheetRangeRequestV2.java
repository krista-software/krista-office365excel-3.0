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

package app.krista.extensions.essentials.spreadsheets.office365excel.graph.v2;

import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.HttpMethod;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.requests.WorkbookWorksheetRangeRequest;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.List;

public class WorkbookWorksheetRangeRequestV2 extends WorkbookWorksheetRangeRequest {
    /**
     * The request for this WorkbookWorksheetUsedRange
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public WorkbookWorksheetRangeRequestV2(@NotNull String requestUrl, @NotNull IBaseClient<?> client,
                                           @Nullable List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    public WorkbookRange patch(@NotNull WorkbookRange workbookRange) throws ClientException {
        return send(HttpMethod.PATCH, workbookRange);
    }
}
