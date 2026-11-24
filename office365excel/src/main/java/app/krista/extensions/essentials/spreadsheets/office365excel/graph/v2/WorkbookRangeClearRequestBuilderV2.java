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

import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.models.WorkbookRangeClearParameterSet;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.requests.WorkbookRangeClearRequestBuilder;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.List;

public class WorkbookRangeClearRequestBuilderV2 extends WorkbookRangeClearRequestBuilder {

    private static final String CLEAR = "clear";

    private WorkbookRangeClearRequestBuilderV2(@NotNull String requestUrl, @NotNull IBaseClient<?> client, @Nullable List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    private WorkbookRangeClearRequestBuilderV2(@NotNull String requestUrl, @NotNull IBaseClient<?> client, @Nullable List<? extends Option> requestOptions, @NotNull WorkbookRangeClearParameterSet parameters) {
        super(requestUrl, client, requestOptions, parameters);
    }

    public static WorkbookRangeClearRequestBuilderV2 newBuilder(@NotNull String requestUrl, @NotNull IBaseClient<?> client, @Nullable List<? extends Option> requestOptions, @NotNull WorkbookRangeClearParameterSet parameters) {
        String requestURL = requestUrl + Constants.BACKWARD_SLASH + CLEAR;
        return new WorkbookRangeClearRequestBuilderV2(requestURL, client, requestOptions, parameters);
    }

    public static WorkbookRangeClearRequestBuilderV2 newBuilder(@NotNull String requestUrl, @NotNull IBaseClient<?> client, @Nullable List<? extends Option> requestOptions) {
        String requestURL = requestUrl + Constants.BACKWARD_SLASH + CLEAR;
        return new WorkbookRangeClearRequestBuilderV2(requestURL, client, requestOptions);
    }

}
