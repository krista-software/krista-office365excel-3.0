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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.attributes;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link Print}.
 */
@DisplayName("Print")
class PrintTest {

    @Test
    @DisplayName("prettyPrint with non-null object does not throw exception")
    void prettyPrintWithNonNullObjectDoesNotThrow() {
        assertDoesNotThrow(() -> Print.prettyPrint("Hello World"));
    }

    @Test
    @DisplayName("prettyPrint with a Map object does not throw exception")
    void prettyPrintWithMapDoesNotThrow() {
        Map<String, Object> testMap = Map.of("key1", "value1", "key2", 42);
        assertDoesNotThrow(() -> Print.prettyPrint(testMap));
    }

    @Test
    @DisplayName("prettyPrint with an integer object does not throw exception")
    void prettyPrintWithIntegerDoesNotThrow() {
        assertDoesNotThrow(() -> Print.prettyPrint(123));
    }

    @Test
    @DisplayName("prettyPrint with null object does not throw exception (logs error)")
    void prettyPrintWithNullDoesNotThrow() {
        assertDoesNotThrow(() -> Print.prettyPrint(null));
    }

    @Test
    @DisplayName("prettyPrint with empty string does not throw exception")
    void prettyPrintWithEmptyStringDoesNotThrow() {
        assertDoesNotThrow(() -> Print.prettyPrint(""));
    }

    @Test
    @DisplayName("prettyPrint with complex nested object does not throw exception")
    void prettyPrintWithNestedObjectDoesNotThrow() {
        Map<String, Object> nested = Map.of("outer", Map.of("inner", "value"));
        assertDoesNotThrow(() -> Print.prettyPrint(nested));
    }
}
