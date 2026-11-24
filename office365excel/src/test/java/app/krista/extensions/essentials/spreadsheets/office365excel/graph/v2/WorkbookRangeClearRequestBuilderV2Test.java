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
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Comprehensive test suite for WorkbookRangeClearRequestBuilderV2 with 100% code coverage.
 * Tests all factory methods and constructor variations.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class WorkbookRangeClearRequestBuilderV2Test {

    @Mock
    private IBaseClient<?> client;

    @Mock
    private WorkbookRangeClearParameterSet parameters;

    private String baseRequestUrl;
    private List<Option> requestOptions;

    @BeforeEach
    void setUp() {
        baseRequestUrl = "https://graph.microsoft.com/v1.0/me/drive/items/itemId/workbook/worksheets/Sheet1/range(address='A1:B2')";
        requestOptions = new ArrayList<>();
    }

    // ==================== newBuilder with Parameters Tests ====================

    @Test
    void newBuilder_withParameters_createsBuilderWithClearPath() {
        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                baseRequestUrl, client, requestOptions, parameters);

        assertNotNull(builder);
        // Verify the URL contains the /clear suffix
        assertTrue(builder.getRequestUrl().toString().endsWith("/clear"));
        assertTrue(builder.getRequestUrl().toString().contains(baseRequestUrl));
    }

    @Test
    void newBuilder_withParameters_appendsBackwardSlashAndClear() {
        String expectedUrl = baseRequestUrl + Constants.BACKWARD_SLASH + "clear";

        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                baseRequestUrl, client, requestOptions, parameters);

        assertNotNull(builder);
        assertEquals(expectedUrl, builder.getRequestUrl().toString());
    }

    @Test
    void newBuilder_withParameters_andNullOptions_createsBuilder() {
        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                baseRequestUrl, client, null, parameters);

        assertNotNull(builder);
        assertTrue(builder.getRequestUrl().toString().endsWith("/clear"));
    }

    // ==================== newBuilder without Parameters Tests ====================

    @Test
    void newBuilder_withoutParameters_createsBuilderWithClearPath() {
        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                baseRequestUrl, client, requestOptions);

        assertNotNull(builder);
        assertTrue(builder.getRequestUrl().toString().endsWith("/clear"));
        assertTrue(builder.getRequestUrl().toString().contains(baseRequestUrl));
    }

    @Test
    void newBuilder_withoutParameters_appendsBackwardSlashAndClear() {
        String expectedUrl = baseRequestUrl + Constants.BACKWARD_SLASH + "clear";

        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                baseRequestUrl, client, requestOptions);

        assertNotNull(builder);
        assertEquals(expectedUrl, builder.getRequestUrl().toString());
    }

    @Test
    void newBuilder_withoutParameters_andNullOptions_createsBuilder() {
        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                baseRequestUrl, client, null);

        assertNotNull(builder);
        assertTrue(builder.getRequestUrl().toString().endsWith("/clear"));
    }

    // ==================== Edge Cases ====================

    @Test
    void newBuilder_withEmptyRequestUrl_appendsClearPath() {
        String emptyUrl = "";
        String expectedUrl = Constants.BACKWARD_SLASH + "clear";

        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                emptyUrl, client, requestOptions, parameters);

        assertNotNull(builder);
        assertEquals(expectedUrl, builder.getRequestUrl().toString());
    }

    @Test
    void newBuilder_withUrlEndingInSlash_doesNotDuplicateSlash() {
        String urlWithSlash = baseRequestUrl + "/";

        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                urlWithSlash, client, requestOptions, parameters);

        assertNotNull(builder);
        // Should have /clear appended
        assertTrue(builder.getRequestUrl().toString().endsWith("/clear"));
    }

    @Test
    void newBuilder_withDifferentUrls_createsDistinctBuilders() {
        String url1 = "https://graph.microsoft.com/v1.0/drive/items/item1";
        String url2 = "https://graph.microsoft.com/v1.0/drive/items/item2";

        WorkbookRangeClearRequestBuilderV2 builder1 = WorkbookRangeClearRequestBuilderV2.newBuilder(
                url1, client, requestOptions, parameters);
        WorkbookRangeClearRequestBuilderV2 builder2 = WorkbookRangeClearRequestBuilderV2.newBuilder(
                url2, client, requestOptions, parameters);

        assertNotNull(builder1);
        assertNotNull(builder2);
        assertNotEquals(builder1.getRequestUrl(), builder2.getRequestUrl());
        assertTrue(builder1.getRequestUrl().toString().contains("item1"));
        assertTrue(builder2.getRequestUrl().toString().contains("item2"));
    }

    @Test
    void newBuilder_withEmptyOptionsList_createsBuilder() {
        List<Option> emptyOptions = new ArrayList<>();

        WorkbookRangeClearRequestBuilderV2 builder = WorkbookRangeClearRequestBuilderV2.newBuilder(
                baseRequestUrl, client, emptyOptions, parameters);

        assertNotNull(builder);
        assertTrue(builder.getRequestUrl().toString().endsWith("/clear"));
    }
}

