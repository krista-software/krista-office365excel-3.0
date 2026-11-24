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

package app.krista.extensions.essentials.spreadsheets.office365excel.data.formatter;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link NumberFormatter}.
 */
@DisplayName("NumberFormatter")
class NumberFormatterTest {

    @Test
    @DisplayName("Constructor stores the format string")
    void constructorStoresFormat() throws Exception {
        NumberFormatter formatter = new NumberFormatter("0.00");
        Field formatField = NumberFormatter.class.getDeclaredField("format");
        formatField.setAccessible(true);
        assertEquals("0.00", formatField.get(formatter));
    }

    @Test
    @DisplayName("format() returns Double.parseDouble of valid integer string")
    void formatReturnsDoubleForValidInteger() {
        NumberFormatter formatter = new NumberFormatter("0");
        Object result = formatter.format("42");
        assertEquals(42.0, result);
        assertInstanceOf(Double.class, result);
    }

    @Test
    @DisplayName("format() returns Double.parseDouble of valid decimal string")
    void formatReturnsDoubleForValidDecimal() {
        NumberFormatter formatter = new NumberFormatter("0.00");
        Object result = formatter.format("3.14159");
        assertEquals(3.14159, result);
    }

    @Test
    @DisplayName("format() returns Double.parseDouble of negative number string")
    void formatReturnsDoubleForNegativeNumber() {
        NumberFormatter formatter = new NumberFormatter("0.00");
        Object result = formatter.format("-99.5");
        assertEquals(-99.5, result);
    }

    @Test
    @DisplayName("format() throws NumberFormatException on invalid string")
    void formatThrowsOnInvalidString() {
        NumberFormatter formatter = new NumberFormatter("0.00");
        assertThrows(NumberFormatException.class, () -> formatter.format("not-a-number"));
    }

    @Test
    @DisplayName("format() throws NumberFormatException on empty string")
    void formatThrowsOnEmptyString() {
        NumberFormatter formatter = new NumberFormatter("0.00");
        assertThrows(NumberFormatException.class, () -> formatter.format(""));
    }

    @Test
    @DisplayName("format() throws NullPointerException on null string")
    void formatThrowsOnNull() {
        NumberFormatter formatter = new NumberFormatter("0.00");
        assertThrows(NullPointerException.class, () -> formatter.format(null));
    }
}
