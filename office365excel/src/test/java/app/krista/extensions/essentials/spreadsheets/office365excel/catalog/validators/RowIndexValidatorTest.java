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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class RowIndexValidatorTest {

    private RowIndexValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new RowIndexValidator();
        context = new HashMap<>();
    }

    // --- validate tests ---

    @Test
    void validate_positiveNumber_returnsTrue() {
        assertTrue(validator.validate("5", context));
    }

    @Test
    void validate_one_returnsTrue() {
        assertTrue(validator.validate("1", context));
    }

    @Test
    void validate_positiveDecimal_returnsTrue() {
        assertTrue(validator.validate("2.5", context));
    }

    @Test
    void validate_zero_returnsFalse() {
        assertFalse(validator.validate("0", context));
    }

    @Test
    void validate_negativeNumber_returnsFalse() {
        assertFalse(validator.validate("-1", context));
    }

    @Test
    void validate_nonNumber_returnsFalse() {
        assertFalse(validator.validate("abc", context));
    }

    @Test
    void validate_nullInput_returnsFalse() {
        assertFalse(validator.validate(null, context));
    }

    @Test
    void validate_emptyInput_returnsFalse() {
        assertFalse(validator.validate("", context));
    }

    @Test
    void validate_whitespaceOnly_returnsFalse() {
        assertFalse(validator.validate("   ", context));
    }

    @Test
    void validate_largePositiveNumber_returnsTrue() {
        assertTrue(validator.validate("1048576", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Row Index", validator.getFetchFieldName());
    }

    @Test
    void getFieldType_returnsNumberFieldLogicalName() {
        assertEquals("com.krista.fields.Number", validator.getFieldType());
    }

    @Test
    void getFetchStepMessage_returnsNonEmpty() {
        String msg = validator.getFetchStepMessage();
        assertNotNull(msg);
        assertFalse(msg.isEmpty());
    }

    @Test
    void getConfirmStepMessage_returnsNonNull() {
        assertNotNull(validator.getConfirmStepMessage("0", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("0"));
    }

    @Test
    void getConfirmStepMessage_containsRowIndex() {
        String msg = validator.getConfirmStepMessage("42", context);
        assertTrue(msg.contains("42"));
    }

    @Test
    void getErrMessage_containsRowIndex() {
        String msg = validator.getErrMessage("42");
        assertTrue(msg.contains("42"));
    }
}
