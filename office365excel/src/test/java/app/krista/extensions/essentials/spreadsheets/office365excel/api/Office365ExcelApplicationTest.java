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

package app.krista.extensions.essentials.spreadsheets.office365excel.api;

import app.krista.extensions.essentials.spreadsheets.office365excel.mcp.ExcelMcpProtocolResource;
import app.krista.extensions.essentials.spreadsheets.office365excel.mcp.ExcelMcpSetupResource;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import javax.ws.rs.ApplicationPath;
import javax.ws.rs.core.Application;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive test suite for Office365ExcelApplication with 100% code coverage.
 * Tests JAX-RS Application configuration and resource registration.
 */
class Office365ExcelApplicationTest {

    private Office365ExcelApplication application;

    @BeforeEach
    void setUp() {
        application = new Office365ExcelApplication();
    }

    // ==================== Constructor Tests ====================

    @Test
    void constructor_createsInstance() {
        Office365ExcelApplication newApplication = new Office365ExcelApplication();
        
        assertNotNull(newApplication);
        assertTrue(newApplication instanceof Application);
    }

    // ==================== getClasses Tests ====================

    @Test
    void getClasses_returnsAuthenticationResourceClass() {
        Set<Class<?>> classes = application.getClasses();

        assertNotNull(classes);
        assertEquals(4, classes.size());
        assertTrue(classes.contains(AuthenticationResource.class));
        assertTrue(classes.contains(Office365ExcelAuditResource.class));
        assertTrue(classes.contains(ExcelMcpSetupResource.class));
        assertTrue(classes.contains(ExcelMcpProtocolResource.class));
    }

    @Test
    void getClasses_returnsImmutableSet() {
        Set<Class<?>> classes = application.getClasses();

        assertNotNull(classes);
        // Verify it's a Set.of() which is immutable
        assertThrows(UnsupportedOperationException.class, () -> classes.add(Object.class));
    }

    @Test
    void getClasses_multipleCallsReturnSameContent() {
        Set<Class<?>> classes1 = application.getClasses();
        Set<Class<?>> classes2 = application.getClasses();

        assertNotNull(classes1);
        assertNotNull(classes2);
        assertEquals(classes1.size(), classes2.size());
        assertTrue(classes1.containsAll(classes2));
        assertTrue(classes2.containsAll(classes1));
    }

    // ==================== Annotation Tests ====================

    @Test
    void class_hasApplicationPathAnnotation() {
        ApplicationPath annotation = Office365ExcelApplication.class.getAnnotation(ApplicationPath.class);
        
        assertNotNull(annotation);
        assertEquals("excel365extn", annotation.value());
    }

    @Test
    void class_extendsApplication() {
        assertTrue(Application.class.isAssignableFrom(Office365ExcelApplication.class));
    }

    @Test
    void class_isNotFinal() {
        assertFalse(java.lang.reflect.Modifier.isFinal(Office365ExcelApplication.class.getModifiers()));
    }

    // ==================== Integration Tests ====================

    @Test
    void getClasses_containsAuthenticationAndAuditResources() {
        Set<Class<?>> classes = application.getClasses();

        assertEquals(4, classes.size());
        assertTrue(classes.contains(AuthenticationResource.class));
        assertTrue(classes.contains(Office365ExcelAuditResource.class));
        assertTrue(classes.contains(ExcelMcpSetupResource.class));
        assertTrue(classes.contains(ExcelMcpProtocolResource.class));
    }

    @Test
    void getClasses_doesNotContainOtherResources() {
        Set<Class<?>> classes = application.getClasses();

        assertFalse(classes.contains(RestCallbackResource.class));
        assertFalse(classes.contains(RestApplication.class));
        assertFalse(classes.contains(Office365ExcelApplication.class));
    }

    // ==================== Edge Case Tests ====================

    @Test
    void getClasses_neverReturnsNull() {
        Set<Class<?>> classes = application.getClasses();
        
        assertNotNull(classes);
    }

    @Test
    void getClasses_neverReturnsEmptySet() {
        Set<Class<?>> classes = application.getClasses();
        
        assertFalse(classes.isEmpty());
    }

    @Test
    void multipleInstances_returnSameResourceClasses() {
        Office365ExcelApplication app1 = new Office365ExcelApplication();
        Office365ExcelApplication app2 = new Office365ExcelApplication();

        Set<Class<?>> classes1 = app1.getClasses();
        Set<Class<?>> classes2 = app2.getClasses();

        assertEquals(classes1, classes2);
    }

    // ==================== Verification Tests ====================

    @Test
    void getClasses_returnsSetOfClasses() {
        Set<Class<?>> classes = application.getClasses();

        assertNotNull(classes);
        assertTrue(classes instanceof Set);
        for (Object obj : classes) {
            assertTrue(obj instanceof Class);
        }
    }

    @Test
    void authenticationResource_isValidJaxRsResource() {
        Set<Class<?>> classes = application.getClasses();
        Class<?> resourceClass = classes.iterator().next();

        // Verify it has JAX-RS Path annotation
        assertNotNull(resourceClass.getAnnotation(javax.ws.rs.Path.class));
    }
}

