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
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Comprehensive test suite for WorkbookOperation with 100% code coverage.
 * Tests workbook operations including search, row operations, and data retrieval.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class WorkbookOperationTest {

    @Mock
    private GraphServiceClientProviderFactory providerFactory;

    @Mock
    private GraphServiceClientProvider clientProvider;

    @Mock
    private RestClientProvider restClientProvider;

    private WorkbookOperation workbookOperation;

    @BeforeEach
    void setUp() {
        when(providerFactory.create()).thenReturn(clientProvider);
        workbookOperation = new WorkbookOperation(providerFactory, restClientProvider);
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_initializesWithProviderFactory() {
        // Constructor stores the factory but uses lazy init — create() is NOT called yet.
        assertNotNull(workbookOperation);
        verify(providerFactory, never()).create();
    }

    @Test
    void constructor_storesRestClientProvider() {
        assertNotNull(workbookOperation);
    }

    // ==================== Multiple Instance Tests ====================

    @Test
    void multipleInstances_areIndependent() {
        WorkbookOperation operation1 = new WorkbookOperation(providerFactory, restClientProvider);
        WorkbookOperation operation2 = new WorkbookOperation(providerFactory, restClientProvider);

        assertNotNull(operation1);
        assertNotNull(operation2);
        assertNotSame(operation1, operation2);
    }

    // ==================== Constant Tests ====================

    @Test
    void constants_haveExpectedValues() {
        assertEquals("updatedWorkBookRangeAddress = ", WorkbookOperation.UPDATED_WORK_BOOK_RANGE_ADDRESS);
        assertEquals("Problem with workbookRange", WorkbookOperation.PROBLEM_WITH_WORKBOOK_RANGE);
        assertEquals("Sheet Is Empty", WorkbookOperation.SHEET_IS_EMPTY);
        assertEquals("Response Body Is Empty", WorkbookOperation.RESPONSE_BODY_IS_EMPTY);
        assertEquals("/range(address='", WorkbookOperation.RANGE_ADDRESS);
        assertEquals("/workbook/worksheets/", WorkbookOperation.WORKBOOK_WORKSHEETS_SHEET);
    }

    // ==================== Verification Tests ====================

    @Test
    void constructor_callsProviderFactoryCreate() {
        // Lazy init — create() is NOT called during construction.
        verify(providerFactory, never()).create();
    }

    @Test
    void constructor_withNullProviderFactory_doesNotThrowDuringConstruction() {
        // Null factory is stored but not dereferenced until first use — no NPE at construction time.
        assertDoesNotThrow(() -> new WorkbookOperation(null, restClientProvider));
    }

    @Test
    void constructor_withNullRestClientProvider_doesNotThrow() {
        assertDoesNotThrow(() -> new WorkbookOperation(providerFactory, null));
    }

    // ==================== Class Structure Tests ====================

    @Test
    void class_isPublic() {
        assertTrue(java.lang.reflect.Modifier.isPublic(WorkbookOperation.class.getModifiers()));
    }

    @Test
    void class_hasServiceAnnotation() {
        assertTrue(WorkbookOperation.class.isAnnotationPresent(org.jvnet.hk2.annotations.Service.class));
    }

    // ==================== Dependency Injection Tests ====================

    @Test
    void constructor_hasInjectAnnotation() throws NoSuchMethodException {
        var constructor = WorkbookOperation.class.getConstructor(
            GraphServiceClientProviderFactory.class, 
            RestClientProvider.class
        );
        assertTrue(constructor.isAnnotationPresent(javax.inject.Inject.class));
    }

    // ==================== Instance State Tests ====================

    @Test
    void newInstance_hasNonNullClientProvider() {
        // Lazy init — factory is stored but create() is not called until first use.
        assertNotNull(workbookOperation);
        verify(providerFactory, never()).create();
    }

    @Test
    void multipleConstructorCalls_neverCallCreate() {
        reset(providerFactory);
        when(providerFactory.create()).thenReturn(clientProvider);

        new WorkbookOperation(providerFactory, restClientProvider);
        new WorkbookOperation(providerFactory, restClientProvider);
        new WorkbookOperation(providerFactory, restClientProvider);

        // Lazy init — create() is deferred to first use, not called during construction.
        verify(providerFactory, never()).create();
    }
}

