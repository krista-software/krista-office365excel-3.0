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
import com.microsoft.graph.options.FunctionOption;
import com.microsoft.graph.options.Option;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.lang.reflect.Field;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Comprehensive test suite for {@link WorkbookWorksheetRangeRequestBuilderV2} with 100% code coverage.
 * <p>
 * Covers:
 * <ul>
 *   <li>Constructor (with populated, empty, and null request options)</li>
 *   <li>{@code buildRequest(Option...)} varargs overload: no-arg, single option, multiple options, {@code null} varargs</li>
 *   <li>{@code buildRequest(List<? extends Option>)} list overload: empty list, populated list</li>
 *   <li>Function-options forwarding in BOTH overloads — ensures the {@code address} / other function parameters
 *       are not dropped (regression coverage for the V2 list-overload bug that silently returned the entire sheet)</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class WorkbookWorksheetRangeRequestBuilderV2Test {

    @Mock
    private IBaseClient<?> client;

    @Mock
    private WorkbookWorksheetRangeParameterSet parameters;

    private String requestUrl;
    private List<Option> requestOptions;
    private WorkbookWorksheetRangeRequestBuilderV2 builder;

    @BeforeEach
    void setUp() {
        requestUrl = "https://graph.microsoft.com/v1.0/me/drive/items/itemId/workbook/worksheets/Sheet1/range";
        requestOptions = new ArrayList<>();
        builder = new WorkbookWorksheetRangeRequestBuilderV2(requestUrl, client, requestOptions, parameters);
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_initializesCorrectly() {
        WorkbookWorksheetRangeRequestBuilderV2 newBuilder = new WorkbookWorksheetRangeRequestBuilderV2(
                requestUrl, client, requestOptions, parameters);

        assertNotNull(newBuilder);
    }

    @Test
    void constructor_withNullOptions_initializesCorrectly() {
        WorkbookWorksheetRangeRequestBuilderV2 newBuilder = new WorkbookWorksheetRangeRequestBuilderV2(
                requestUrl, client, null, parameters);

        assertNotNull(newBuilder);
    }

    @Test
    void constructor_withPopulatedOptions_initializesCorrectly() {
        List<Option> populated = new ArrayList<>();
        populated.add(mock(Option.class));

        WorkbookWorksheetRangeRequestBuilderV2 newBuilder = new WorkbookWorksheetRangeRequestBuilderV2(
                requestUrl, client, populated, parameters);

        assertNotNull(newBuilder);
    }

    @Test
    void constructor_withDifferentUrls_createsDistinctBuilders() {
        String url1 = "https://graph.microsoft.com/v1.0/drive/items/item1";
        String url2 = "https://graph.microsoft.com/v1.0/drive/items/item2";

        WorkbookWorksheetRangeRequestBuilderV2 builder1 = new WorkbookWorksheetRangeRequestBuilderV2(
                url1, client, requestOptions, parameters);
        WorkbookWorksheetRangeRequestBuilderV2 builder2 = new WorkbookWorksheetRangeRequestBuilderV2(
                url2, client, requestOptions, parameters);

        assertNotNull(builder1);
        assertNotNull(builder2);
        assertNotEquals(builder1.getRequestUrl(), builder2.getRequestUrl());
    }

    // ==================== buildRequest(Option...) Varargs Tests ====================

    @Test
    void buildRequest_varargs_withNoOptions_returnsRequestV2() {
        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest();

        assertNotNull(request);
    }

    @Test
    void buildRequest_varargs_withSingleOption_returnsRequestV2() {
        Option option = mock(Option.class);

        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(option);

        assertNotNull(request);
    }

    @Test
    void buildRequest_varargs_withMultipleOptions_returnsRequestV2() {
        Option option1 = mock(Option.class);
        Option option2 = mock(Option.class);

        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(option1, option2);

        assertNotNull(request);
    }

    @Test
    void buildRequest_varargs_withNullVarargs_returnsRequestV2() {
        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest((Option[]) null);

        assertNotNull(request);
    }

    // ==================== buildRequest(List) Tests ====================

    @Test
    void buildRequest_list_withEmptyList_returnsRequestV2() {
        List<Option> emptyOptions = new ArrayList<>();

        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(emptyOptions);

        assertNotNull(request);
    }

    @Test
    void buildRequest_list_withOptions_returnsRequestV2() {
        List<Option> options = new ArrayList<>();
        options.add(mock(Option.class));
        options.add(mock(Option.class));

        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(options);

        assertNotNull(request);
    }

    @Test
    void buildRequest_list_setsCorrectUrl() throws Exception {
        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(new ArrayList<>());

        assertNotNull(request.getRequestUrl());
        assertEquals(new URL(requestUrl), request.getRequestUrl());
    }

    @Test
    void buildRequest_list_setsCorrectClient() {
        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(new ArrayList<>());

        assertSame(client, request.getClient());
    }

    // ==================== Function-Options Forwarding Tests ====================
    // These tests cover the for-loop bodies (lines 23-25 and 38-40 of the class under test).
    // They verify that function options such as `address='A1:B2'` are copied to the returned
    // request on BOTH the varargs AND list overloads — the regression-coverage tests for the
    // bug where the list overload was silently dropping the range address.

    @Test
    void buildRequest_varargs_forwardsFunctionOptions() throws Exception {
        FunctionOption addressOption = new FunctionOption("address", "A1:B2");
        FunctionOption valuesOnlyOption = new FunctionOption("valuesOnly", Boolean.TRUE);
        injectFunctionOptions(builder, Arrays.asList(addressOption, valuesOnlyOption));

        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest();

        assertNotNull(request);
        List<FunctionOption> forwarded = request.getFunctionOptions();
        assertNotNull(forwarded);
        assertTrue(forwarded.contains(addressOption),
                "Varargs overload must forward the address function option");
        assertTrue(forwarded.contains(valuesOnlyOption),
                "Varargs overload must forward all function options");
    }

    @Test
    void buildRequest_list_forwardsFunctionOptions() throws Exception {
        FunctionOption addressOption = new FunctionOption("address", "A9:AM42");
        injectFunctionOptions(builder, Collections.singletonList(addressOption));

        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(new ArrayList<>());

        assertNotNull(request);
        List<FunctionOption> forwarded = request.getFunctionOptions();
        assertNotNull(forwarded);
        assertTrue(forwarded.contains(addressOption),
                "List overload must forward the address function option "
                        + "(regression: V2 list-overload previously dropped functionOptions)");
    }

    @Test
    void buildRequest_list_withSessionHeader_stillForwardsFunctionOptions() throws Exception {
        // Mirrors production call pattern: list containing a session-header option plus
        // function options carrying the range address. Both MUST reach the request.
        FunctionOption addressOption = new FunctionOption("address", "Sheet1!A1:D10");
        injectFunctionOptions(builder, Collections.singletonList(addressOption));

        List<Option> sessionOptions = new ArrayList<>();
        sessionOptions.add(mock(Option.class));

        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(sessionOptions);

        assertNotNull(request);
        assertTrue(request.getFunctionOptions().contains(addressOption));
    }

    @Test
    void buildRequest_varargs_withEmptyFunctionOptions_forwardsNone() {
        // functionOptions is empty by default (mock parameters provide none); the for-loop
        // body must be skipped without adding anything.
        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest();

        assertNotNull(request);
        // A fresh mock-backed builder should yield a request with no function options added.
        assertTrue(request.getFunctionOptions() == null || request.getFunctionOptions().isEmpty());
    }

    @Test
    void buildRequest_list_withEmptyFunctionOptions_forwardsNone() {
        WorkbookWorksheetRangeRequestV2 request = builder.buildRequest(new ArrayList<>());

        assertNotNull(request);
        assertTrue(request.getFunctionOptions() == null || request.getFunctionOptions().isEmpty());
    }

    // ==================== Behavioral Consistency Tests ====================

    @Test
    void buildRequest_calledMultipleTimes_createsDistinctRequests() {
        WorkbookWorksheetRangeRequestV2 request1 = builder.buildRequest();
        WorkbookWorksheetRangeRequestV2 request2 = builder.buildRequest();

        assertNotNull(request1);
        assertNotNull(request2);
        assertNotSame(request1, request2);
    }

    @Test
    void buildRequest_varargs_delegatesToListOverload() {
        // The varargs overload must produce a request that behaves identically to the
        // list overload (same URL, same client).
        WorkbookWorksheetRangeRequestV2 viaVarargs = builder.buildRequest();
        WorkbookWorksheetRangeRequestV2 viaList = builder.buildRequest(new ArrayList<>());

        assertEquals(viaVarargs.getRequestUrl(), viaList.getRequestUrl());
        assertSame(viaVarargs.getClient(), viaList.getClient());
    }

    @Test
    void buildRequest_preservesBuilderState() {
        WorkbookWorksheetRangeRequestV2 request1 = builder.buildRequest();
        WorkbookWorksheetRangeRequestV2 request2 = builder.buildRequest();

        assertEquals(request1.getRequestUrl(), request2.getRequestUrl());
        assertSame(request1.getClient(), request2.getClient());
    }

    @Test
    void buildRequest_withDifferentOptionCombinations_allSucceed() {
        Option option1 = mock(Option.class);
        Option option2 = mock(Option.class);

        assertNotNull(builder.buildRequest());
        assertNotNull(builder.buildRequest(option1));
        assertNotNull(builder.buildRequest(option1, option2));
        assertNotNull(builder.buildRequest(new ArrayList<>()));
        assertNotNull(builder.buildRequest(Arrays.asList(option1, option2)));
    }

    // ==================== Helpers ====================

    /**
     * Injects function options into the protected {@code functionOptions} field inherited from
     * {@code BaseFunctionRequestBuilder}. Required because {@code WorkbookWorksheetRangeParameterSet}
     * is mocked (so {@code getFunctionOptions()} returns an empty list by default) and there is no
     * public setter.
     */
    private static void injectFunctionOptions(WorkbookWorksheetRangeRequestBuilderV2 target,
                                              List<FunctionOption> options) throws Exception {
        Field field = findField(target.getClass(), "functionOptions");
        field.setAccessible(true);
        @SuppressWarnings("unchecked")
        List<FunctionOption> existing = (List<FunctionOption>) field.get(target);
        if (existing == null) {
            field.set(target, new ArrayList<>(options));
        } else {
            existing.clear();
            existing.addAll(options);
        }
    }

    private static Field findField(Class<?> clazz, String name) throws NoSuchFieldException {
        Class<?> current = clazz;
        while (current != null) {
            try {
                return current.getDeclaredField(name);
            } catch (NoSuchFieldException ignored) {
                current = current.getSuperclass();
            }
        }
        throw new NoSuchFieldException(name);
    }
}
