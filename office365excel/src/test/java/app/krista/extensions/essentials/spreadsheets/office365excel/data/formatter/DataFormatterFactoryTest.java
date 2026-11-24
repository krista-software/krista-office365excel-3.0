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

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link DataFormatterFactory#create(String, String)}.
 */
@DisplayName("DataFormatterFactory - create()")
class DataFormatterFactoryTest {

    @Test
    @DisplayName("valueType=Double with format containing '-' returns DateFormatter")
    void doubleWithDashReturnsDateFormatter() {
        DataFormatter formatter = DataFormatterFactory.create("Double", "MM-dd-yyyy");
        assertInstanceOf(DateFormatter.class, formatter);
    }

    @Test
    @DisplayName("valueType=Double with format containing '/' returns DateFormatter")
    void doubleWithSlashReturnsDateFormatter() {
        DataFormatter formatter = DataFormatterFactory.create("Double", "MM/dd/yyyy");
        assertInstanceOf(DateFormatter.class, formatter);
    }

    @Test
    @DisplayName("valueType=Double with format containing both '-' and '/' returns DateFormatter")
    void doubleWithDashAndSlashReturnsDateFormatter() {
        DataFormatter formatter = DataFormatterFactory.create("Double", "dd-MM/yyyy");
        assertInstanceOf(DateFormatter.class, formatter);
    }

    @Test
    @DisplayName("valueType=Double with numeric format returns NumberFormatter")
    void doubleWithNumericFormatReturnsNumberFormatter() {
        DataFormatter formatter = DataFormatterFactory.create("Double", "0.00");
        assertInstanceOf(NumberFormatter.class, formatter);
    }

    @Test
    @DisplayName("valueType=String returns StringFormatter")
    void stringTypeReturnsStringFormatter() {
        DataFormatter formatter = DataFormatterFactory.create("String", "General");
        assertInstanceOf(StringFormatter.class, formatter);
    }

    @Test
    @DisplayName("Other valueType returns StringFormatter")
    void otherTypeReturnsStringFormatter() {
        DataFormatter formatter = DataFormatterFactory.create("Boolean", "General");
        assertInstanceOf(StringFormatter.class, formatter);
    }

    @Test
    @DisplayName("null valueType returns StringFormatter")
    void nullValueTypeReturnsStringFormatter() {
        DataFormatter formatter = DataFormatterFactory.create(null, "someFormat");
        assertInstanceOf(StringFormatter.class, formatter);
    }
}
