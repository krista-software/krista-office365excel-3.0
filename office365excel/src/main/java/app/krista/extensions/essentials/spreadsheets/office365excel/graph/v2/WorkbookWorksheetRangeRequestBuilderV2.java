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

import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.models.WorkbookWorksheetRangeParameterSet;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.requests.WorkbookWorksheetRangeRequestBuilder;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.List;

public class WorkbookWorksheetRangeRequestBuilderV2 extends WorkbookWorksheetRangeRequestBuilder {

    public WorkbookWorksheetRangeRequestBuilderV2(@NotNull String requestUrl, @NotNull IBaseClient<?> client,
                                                  @Nullable List<? extends Option> requestOptions,
                                                  @NotNull WorkbookWorksheetRangeParameterSet parameters) {
        super(requestUrl, client, requestOptions, parameters);
    }

    @Override
    public WorkbookWorksheetRangeRequestV2 buildRequest(@Nullable final Option... requestOptions) {
        return buildRequest(getOptions(requestOptions));
    }

    @Override
    public WorkbookWorksheetRangeRequestV2 buildRequest(
            @NotNull final List<? extends Option> requestOptions) {
        WorkbookWorksheetRangeRequestV2 request = new WorkbookWorksheetRangeRequestV2(
                getRequestUrl(),
                getClient(),
                requestOptions);
        // Parent's varargs overload copies function options (e.g. address=A9:AM42);
        // the list overload must do the same or the range-function parameters are lost.
        for (com.microsoft.graph.options.FunctionOption option : functionOptions) {
            request.addFunctionOption(option);
        }
        return request;
    }
}
