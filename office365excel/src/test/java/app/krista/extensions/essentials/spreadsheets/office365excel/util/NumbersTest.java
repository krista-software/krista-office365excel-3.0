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

package app.krista.extensions.essentials.spreadsheets.office365excel.util;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for Numbers utility class
 */
@DisplayName("Numbers Utility Tests")
class NumbersTest {

    @ParameterizedTest
    @ValueSource(strings = {
        "123",
        "123.45",
        "-123",
        "+123",
        "0.5",
        ".5",
        "1,234",
        "1,234.56",
        "1,234,567.89",
        "1.23e10",
        "1.23E10",
        "1.23e-10",
        "-1.23e+10",
        "0",
        "0.0",
        "  123.45  "
    })
    @DisplayName("isNumber - Valid number formats")
    void testIsNumber_ValidNumbers(String input) {
        assertTrue(Numbers.isNumber(input), "Should recognize '" + input + "' as a valid number");
    }

    @ParameterizedTest
    @ValueSource(strings = {
        "abc",
        "12.34.56",
        "1.2.3",
        "",
        "  ",
        "12a34",
        "12 34",
        "e10",
        "1.23ee10",
        "null",
        "NaN",
        "Infinity"
    })
    @DisplayName("isNumber - Invalid number formats")
    void testIsNumber_InvalidNumbers(String input) {
        assertFalse(Numbers.isNumber(input), "Should NOT recognize '" + input + "' as a valid number");
    }

    @ParameterizedTest
    @ValueSource(strings = {
        "1,23",      // Becomes "123" after comma removal - VALID
        "1,,234"     // Becomes "1234" after comma removal - VALID
    })
    @DisplayName("isNumber - Valid numbers with malformed commas (commas are stripped)")
    void testIsNumber_MalformedCommasAreValid(String input) {
        // Note: The implementation removes ALL commas before validation,
        // so even malformed comma placement results in valid numbers
        assertTrue(Numbers.isNumber(input), "Should recognize '" + input + "' as valid (commas are stripped)");
    }
}

