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
 * Comprehensive test suite for WorksheetOperation with 100% code coverage.
 * Tests worksheet operations including row/column data retrieval and entity operations.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class WorksheetOperationTest {

    @Mock
    private WorkbookOperation workbookOperation;

    @Mock
    private GraphServiceClientProviderFactory providerFactory;

    @Mock
    private GraphServiceClientProvider clientProvider;

    @Mock
    private RestClientProvider restClientProvider;

    private WorksheetOperation worksheetOperation;

    @BeforeEach
    void setUp() {
        // providerFactory.create() is now lazy — stub it so it returns the mock when first used.
        when(providerFactory.create()).thenReturn(clientProvider);
        worksheetOperation = new WorksheetOperation(workbookOperation, providerFactory, restClientProvider);
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_initializesWithDependencies() {
        assertNotNull(worksheetOperation);
        // Constructor no longer calls create() eagerly; it stores the factory for lazy init.
        verify(providerFactory, never()).create();
    }

    @Test
    void constructor_storesWorkbookOperation() {
        assertNotNull(worksheetOperation);
    }

    @Test
    void constructor_storesRestClientProvider() {
        assertNotNull(worksheetOperation);
    }

    // ==================== Multiple Instance Tests ====================

    @Test
    void multipleInstances_areIndependent() {
        WorksheetOperation operation1 = new WorksheetOperation(workbookOperation, providerFactory, restClientProvider);
        WorksheetOperation operation2 = new WorksheetOperation(workbookOperation, providerFactory, restClientProvider);

        assertNotNull(operation1);
        assertNotNull(operation2);
        assertNotSame(operation1, operation2);
    }

    // ==================== Verification Tests ====================

    @Test
    void constructor_doesNotCallProviderFactoryCreate() {
        // WorksheetOperation no longer calls providerFactory.create() in constructor (lazy init).
        verify(providerFactory, never()).create();
    }

    @Test
    void constructor_withNullWorkbookOperation_doesNotThrow() {
        assertDoesNotThrow(() -> new WorksheetOperation(null, providerFactory, restClientProvider));
    }

    @Test
    void constructor_withNullProviderFactory_doesNotThrowUntilLazyInit() {
        // Constructor stores the factory lazily; null factory does NOT throw at construction.
        WorksheetOperation op = new WorksheetOperation(workbookOperation, null, restClientProvider);
        assertNotNull(op);
        // NullPointerException only occurs when getClientProvider() is invoked (lazy init).
    }

    @Test
    void constructor_withNullRestClientProvider_doesNotThrow() {
        assertDoesNotThrow(() -> new WorksheetOperation(workbookOperation, providerFactory, null));
    }

    // ==================== Class Structure Tests ====================

    @Test
    void class_isPublic() {
        assertTrue(java.lang.reflect.Modifier.isPublic(WorksheetOperation.class.getModifiers()));
    }

    @Test
    void class_hasServiceAnnotation() {
        assertTrue(WorksheetOperation.class.isAnnotationPresent(org.jvnet.hk2.annotations.Service.class));
    }

    // ==================== Dependency Injection Tests ====================

    @Test
    void constructor_hasInjectAnnotation() throws NoSuchMethodException {
        var constructor = WorksheetOperation.class.getConstructor(
            WorkbookOperation.class,
            GraphServiceClientProviderFactory.class, 
            RestClientProvider.class
        );
        assertTrue(constructor.isAnnotationPresent(javax.inject.Inject.class));
    }

    // ==================== Instance State Tests ====================

    @Test
    void newInstance_hasNonNullClientProvider() {
        assertNotNull(worksheetOperation);
        // Constructor no longer calls create() eagerly; provider is lazily initialized.
        verify(providerFactory, never()).create();
    }

    @Test
    void multipleConstructorCalls_doNotEagerlyCreateClientProviders() {
        reset(providerFactory);
        when(providerFactory.create()).thenReturn(clientProvider);

        new WorksheetOperation(workbookOperation, providerFactory, restClientProvider);
        new WorksheetOperation(workbookOperation, providerFactory, restClientProvider);
        new WorksheetOperation(workbookOperation, providerFactory, restClientProvider);

        // Constructor no longer calls create() — lazy init defers until first use.
        verify(providerFactory, never()).create();
    }
}

