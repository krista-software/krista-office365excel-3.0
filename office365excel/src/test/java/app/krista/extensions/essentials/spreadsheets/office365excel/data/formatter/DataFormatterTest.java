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

import java.text.ParseException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link DataFormatter} interface default method.
 */
@DisplayName("DataFormatter - default format()")
class DataFormatterTest {

    @Test
    @DisplayName("format() returns the data string as-is by default")
    void formatReturnsDataAsIs() throws ParseException {
        DataFormatter formatter = new DataFormatter() {};
        Object result = formatter.format("hello world");
        assertEquals("hello world", result);
    }

    @Test
    @DisplayName("format() returns null string as-is")
    void formatReturnsNullStringAsIs() throws ParseException {
        DataFormatter formatter = new DataFormatter() {};
        Object result = formatter.format(null);
        assertNull(result);
    }

    @Test
    @DisplayName("format() returns empty string as-is")
    void formatReturnsEmptyStringAsIs() throws ParseException {
        DataFormatter formatter = new DataFormatter() {};
        Object result = formatter.format("");
        assertEquals("", result);
    }
}
