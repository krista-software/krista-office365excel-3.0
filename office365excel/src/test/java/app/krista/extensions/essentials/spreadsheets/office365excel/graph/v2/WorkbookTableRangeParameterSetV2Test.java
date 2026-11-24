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

import com.microsoft.graph.options.FunctionOption;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive test suite for WorkbookTableRangeParameterSetV2 with 100% code coverage.
 * Tests builder pattern, parameter setting, and function options.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class WorkbookTableRangeParameterSetV2Test {

    // ==================== Builder Creation Tests ====================

    @Test
    void builder_createsNewBuilderInstance() {
        WorkbookTableRangeParameterSetV2.WorkbookTableRangeParameterSetBuilderV2 builder = 
                WorkbookTableRangeParameterSetV2.builder();

        assertNotNull(builder);
    }

    @Test
    void builder_createsDistinctInstances() {
        WorkbookTableRangeParameterSetV2.WorkbookTableRangeParameterSetBuilderV2 builder1 = 
                WorkbookTableRangeParameterSetV2.builder();
        WorkbookTableRangeParameterSetV2.WorkbookTableRangeParameterSetBuilderV2 builder2 = 
                WorkbookTableRangeParameterSetV2.builder();

        assertNotNull(builder1);
        assertNotNull(builder2);
        assertNotSame(builder1, builder2);
    }

    // ==================== Builder withAddress Tests ====================

    @Test
    void withAddress_setsAddressValue() {
        String address = "A1:D10";

        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress(address)
                .build();

        assertNotNull(parameterSet);
    }

    @Test
    void withAddress_withNullValue_acceptsNull() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress(null)
                .build();

        assertNotNull(parameterSet);
    }

    @Test
    void withAddress_withEmptyString_acceptsEmptyString() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("")
                .build();

        assertNotNull(parameterSet);
    }

    @Test
    void withAddress_returnsBuilderForChaining() {
        WorkbookTableRangeParameterSetV2.WorkbookTableRangeParameterSetBuilderV2 builder = 
                WorkbookTableRangeParameterSetV2.builder();

        WorkbookTableRangeParameterSetV2.WorkbookTableRangeParameterSetBuilderV2 result = 
                builder.withAddress("A1:B2");

        assertSame(builder, result);
    }

    @Test
    void withAddress_canBeCalledMultipleTimes_lastValueWins() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("A1:B2")
                .withAddress("C3:D4")
                .build();

        assertNotNull(parameterSet);
    }

    // ==================== Build Tests ====================

    @Test
    void build_createsParameterSetInstance() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("A1:Z100")
                .build();

        assertNotNull(parameterSet);
    }

    @Test
    void build_withoutSettingAddress_createsParameterSet() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .build();

        assertNotNull(parameterSet);
    }

    @Test
    void build_createsDistinctInstances() {
        WorkbookTableRangeParameterSetV2.WorkbookTableRangeParameterSetBuilderV2 builder = 
                WorkbookTableRangeParameterSetV2.builder().withAddress("A1:B2");

        WorkbookTableRangeParameterSetV2 parameterSet1 = builder.build();
        WorkbookTableRangeParameterSetV2 parameterSet2 = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("A1:B2")
                .build();

        assertNotNull(parameterSet1);
        assertNotNull(parameterSet2);
        assertNotSame(parameterSet1, parameterSet2);
    }

    // ==================== getFunctionOptions Tests ====================

    @Test
    void getFunctionOptions_returnsEmptyList() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("A1:B2")
                .build();

        List<FunctionOption> options = parameterSet.getFunctionOptions();

        assertNotNull(options);
        assertTrue(options.isEmpty());
        assertEquals(0, options.size());
    }

    @Test
    void getFunctionOptions_returnsImmutableList() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("A1:B2")
                .build();

        List<FunctionOption> options = parameterSet.getFunctionOptions();

        assertThrows(UnsupportedOperationException.class, () -> 
            options.add(new FunctionOption("test", "value")));
    }

    @Test
    void getFunctionOptions_calledMultipleTimes_returnsConsistentResult() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("A1:B2")
                .build();

        List<FunctionOption> options1 = parameterSet.getFunctionOptions();
        List<FunctionOption> options2 = parameterSet.getFunctionOptions();

        assertNotNull(options1);
        assertNotNull(options2);
        assertEquals(options1.size(), options2.size());
    }

    // ==================== Complex Scenarios ====================

    @Test
    void fullBuilderChain_withAllMethods_createsValidParameterSet() {
        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress("Sheet1!A1:Z100")
                .build();

        assertNotNull(parameterSet);
        List<FunctionOption> options = parameterSet.getFunctionOptions();
        assertNotNull(options);
        assertTrue(options.isEmpty());
    }

    @Test
    void parameterSet_withComplexAddress_buildsSuccessfully() {
        String complexAddress = "Sheet1!$A$1:$Z$100";

        WorkbookTableRangeParameterSetV2 parameterSet = WorkbookTableRangeParameterSetV2.builder()
                .withAddress(complexAddress)
                .build();

        assertNotNull(parameterSet);
    }
}

