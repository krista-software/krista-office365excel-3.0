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

import com.microsoft.graph.models.WorkbookTableRangeParameterSet;
import com.microsoft.graph.options.FunctionOption;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.List;


public class WorkbookTableRangeParameterSetV2 extends WorkbookTableRangeParameterSet {
    private final String address;

    protected WorkbookTableRangeParameterSetV2(@NotNull WorkbookTableRangeParameterSetBuilderV2 builder) {
        this.address = builder.address;
    }

    @NotNull
    public static WorkbookTableRangeParameterSetBuilderV2 builder() {
        return new WorkbookTableRangeParameterSetBuilderV2();
    }

    public static final class WorkbookTableRangeParameterSetBuilderV2 {

        @Nullable
        protected String address;

        protected WorkbookTableRangeParameterSetBuilderV2() {
        }

        @NotNull
        public WorkbookTableRangeParameterSetBuilderV2 withAddress(@Nullable final String val) {
            this.address = val;
            return this;
        }

        @NotNull
        public WorkbookTableRangeParameterSetV2 build() {
            return new WorkbookTableRangeParameterSetV2(this);
        }
    }

    @NotNull
    public List<FunctionOption> getFunctionOptions() {
        return List.of();
    }
}
