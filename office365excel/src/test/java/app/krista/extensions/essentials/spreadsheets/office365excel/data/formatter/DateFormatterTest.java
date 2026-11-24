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
 * Unit tests for {@link DateFormatter}.
 */
@DisplayName("DateFormatter")
class DateFormatterTest {

    @Test
    @DisplayName("Constructor replaces 'm' with 'M' in format string")
    void constructorReplacesLowercaseMWithUppercase() throws Exception {
        DateFormatter formatter = new DateFormatter("mm-dd-yyyy");
        Field formatField = DateFormatter.class.getDeclaredField("format");
        formatField.setAccessible(true);
        String storedFormat = (String) formatField.get(formatter);
        assertEquals("MM-dd-yyyy", storedFormat);
    }

    @Test
    @DisplayName("Constructor handles format with no 'm' characters")
    void constructorHandlesNoLowercaseM() throws Exception {
        DateFormatter formatter = new DateFormatter("dd/yyyy");
        Field formatField = DateFormatter.class.getDeclaredField("format");
        formatField.setAccessible(true);
        String storedFormat = (String) formatField.get(formatter);
        assertEquals("dd/yyyy", storedFormat);
    }

    @Test
    @DisplayName("Constructor handles format with uppercase 'M' already present")
    void constructorHandlesUppercaseMAlreadyPresent() throws Exception {
        DateFormatter formatter = new DateFormatter("MM/dd/yyyy");
        Field formatField = DateFormatter.class.getDeclaredField("format");
        formatField.setAccessible(true);
        String storedFormat = (String) formatField.get(formatter);
        // "MM" stays "MM" since replace("m","M") only affects lowercase 'm'
        assertEquals("MM/dd/yyyy", storedFormat);
    }

    @Test
    @DisplayName("format() returns data as-is (TODO in code)")
    void formatReturnsDataAsIs() throws ParseException {
        DateFormatter formatter = new DateFormatter("mm-dd-yyyy");
        Object result = formatter.format("12/31/2023");
        assertEquals("12/31/2023", result);
    }

    @Test
    @DisplayName("format() returns null data as-is")
    void formatReturnsNullAsIs() throws ParseException {
        DateFormatter formatter = new DateFormatter("mm-dd-yyyy");
        Object result = formatter.format(null);
        assertNull(result);
    }

    @Test
    @DisplayName("format() returns empty string as-is")
    void formatReturnsEmptyStringAsIs() throws ParseException {
        DateFormatter formatter = new DateFormatter("mm-dd-yyyy");
        Object result = formatter.format("");
        assertEquals("", result);
    }
}
