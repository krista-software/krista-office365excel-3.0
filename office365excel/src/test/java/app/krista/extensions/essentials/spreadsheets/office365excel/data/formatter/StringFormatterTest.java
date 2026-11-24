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
import java.text.ParseException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link StringFormatter}.
 */
@DisplayName("StringFormatter")
class StringFormatterTest {

    @Test
    @DisplayName("Constructor stores the format string in the format field")
    void constructorStoresFormat() throws Exception {
        StringFormatter formatter = new StringFormatter("General");
        Field formatField = StringFormatter.class.getDeclaredField("format");
        formatField.setAccessible(true);
        assertEquals("General", formatField.get(formatter));
    }

    @Test
    @DisplayName("Constructor stores null format")
    void constructorStoresNullFormat() throws Exception {
        StringFormatter formatter = new StringFormatter(null);
        Field formatField = StringFormatter.class.getDeclaredField("format");
        formatField.setAccessible(true);
        assertNull(formatField.get(formatter));
    }

    @Test
    @DisplayName("format() uses default DataFormatter behavior, returning data as-is")
    void formatReturnsDataAsIsUsingDefault() throws ParseException {
        StringFormatter formatter = new StringFormatter("General");
        Object result = formatter.format("hello");
        assertEquals("hello", result);
    }

    @Test
    @DisplayName("format() returns null for null input using default behavior")
    void formatReturnsNullForNullInput() throws ParseException {
        StringFormatter formatter = new StringFormatter("General");
        Object result = formatter.format(null);
        assertNull(result);
    }

    @Test
    @DisplayName("StringFormatter implements DataFormatter interface")
    void implementsDataFormatterInterface() {
        StringFormatter formatter = new StringFormatter("General");
        assertInstanceOf(DataFormatter.class, formatter);
    }
}
