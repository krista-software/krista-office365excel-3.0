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

package app.krista.extensions.essentials.spreadsheets.office365excel.mcp;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extension.impl.anno.Attribute;
import app.krista.extension.impl.anno.Field;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.lang.reflect.Constructor;
import java.lang.reflect.Method;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Tests for {@link MCPArea} — 100 % coverage of constructor and askAgent method.
 */
@ExtendWith(MockitoExtension.class)
class MCPAreaTest {

    @Mock
    private ExcelMcpQueryService mcpQueryService;

    private MCPArea mcpArea;

    @BeforeEach
    void setUp() {
        mcpArea = new MCPArea(mcpQueryService);
    }

    // ── Constructor ────────────────────────────────────────────────────────

    @Test
    void constructor_storesQueryService() {
        assertNotNull(mcpArea, "MCPArea should be instantiated via constructor injection");
    }

    @Test
    void constructor_hasInjectAnnotation() throws NoSuchMethodException {
        Constructor<?> ctor = MCPArea.class.getConstructor(ExcelMcpQueryService.class);
        assertNotNull(ctor.getAnnotation(javax.inject.Inject.class),
                "Constructor must be annotated with @Inject");
    }

    // ── askAgent() ─────────────────────────────────────────────────────────

    @Test
    void askAgent_delegatesToQueryService() {
        ExtensionResponse expected = mock(ExtensionResponse.class);
        when(mcpQueryService.query("list files")).thenReturn(expected);

        ExtensionResponse actual = mcpArea.askAgent("list files");

        assertSame(expected, actual, "askAgent must return exactly what mcpQueryService.query() returns");
        verify(mcpQueryService, times(1)).query("list files");
    }

    @Test
    void askAgent_passesQueryUnmodified() {
        String query = "  some whitespace query  ";
        when(mcpQueryService.query(query)).thenReturn(mock(ExtensionResponse.class));

        mcpArea.askAgent(query);

        verify(mcpQueryService).query(query);
    }

    @Test
    void askAgent_nullQuery_delegatesToQueryService() {
        when(mcpQueryService.query(null)).thenReturn(mock(ExtensionResponse.class));

        ExtensionResponse response = mcpArea.askAgent(null);

        assertNotNull(response);
        verify(mcpQueryService).query(null);
    }

    @Test
    void askAgent_emptyQuery_delegatesToQueryService() {
        when(mcpQueryService.query("")).thenReturn(mock(ExtensionResponse.class));

        ExtensionResponse response = mcpArea.askAgent("");

        assertNotNull(response);
        verify(mcpQueryService).query("");
    }

    // ── Annotation presence on askAgent method ─────────────────────────────

    @Test
    void askAgentMethod_hasFieldDescAnnotation() throws NoSuchMethodException {
        Method method = MCPArea.class.getMethod("askAgent", String.class);
        Field.Desc desc = method.getAnnotation(Field.Desc.class);
        assertNotNull(desc, "@Field.Desc must be present on askAgent");
        assertEquals("Response", desc.name());
        assertEquals("[ FreeForm ]", desc.type());
        assertFalse(desc.required(), "Response field should not be required");
    }

    @Test
    void askAgentParameter_hasFieldAnnotation() throws NoSuchMethodException {
        Method method = MCPArea.class.getMethod("askAgent", String.class);
        Field field = method.getParameters()[0].getAnnotation(Field.class);
        assertNotNull(field, "@Field must be present on the query parameter");
        assertEquals("Query", field.name());
        assertEquals("Paragraph", field.type());
        assertTrue(field.required(), "Query parameter must be required");
    }

    @Test
    void askAgentParameter_hasAttributeVisualWidthL() throws NoSuchMethodException {
        Method method = MCPArea.class.getMethod("askAgent", String.class);
        Field field = method.getParameters()[0].getAnnotation(Field.class);
        Attribute[] attrs = field.attributes();
        assertEquals(1, attrs.length, "Should have exactly one attribute");
        assertEquals("visualWidth", attrs[0].name());
        assertEquals("L", attrs[0].value());
    }
}
