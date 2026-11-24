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

class LookupValueValidatorTest {

    private LookupValueValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new LookupValueValidator();
        context = new HashMap<>();
    }

    // --- validate tests ---

    @Test
    void validate_validValue_returnsTrue() {
        assertTrue(validator.validate("SearchTerm", context));
    }

    @Test
    void validate_singleCharacter_returnsTrue() {
        assertTrue(validator.validate("x", context));
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
    void validate_exactlyMaxLength_returnsTrue() {
        String value = "a".repeat(255);
        assertTrue(validator.validate(value, context));
    }

    @Test
    void validate_exceedsMaxLength_returnsFalse() {
        String value = "a".repeat(256);
        assertFalse(validator.validate(value, context));
    }

    // --- getConfirmStepMessage branching ---

    @Test
    void getConfirmStepMessage_nullValue_returnsEmptyMessage() {
        String msg = validator.getConfirmStepMessage(null, context);
        assertNotNull(msg);
        assertTrue(msg.contains("empty"));
    }

    @Test
    void getConfirmStepMessage_emptyValue_returnsEmptyMessage() {
        String msg = validator.getConfirmStepMessage("", context);
        assertNotNull(msg);
        assertTrue(msg.contains("empty"));
    }

    @Test
    void getConfirmStepMessage_whitespaceValue_returnsEmptyMessage() {
        String msg = validator.getConfirmStepMessage("   ", context);
        assertNotNull(msg);
        assertTrue(msg.contains("empty"));
    }

    @Test
    void getConfirmStepMessage_nonEmptyValue_returnsTooLongMessage() {
        String msg = validator.getConfirmStepMessage("validValue", context);
        assertNotNull(msg);
        assertTrue(msg.contains("too long"));
    }

    // --- getErrMessage branching ---

    @Test
    void getErrMessage_nullValue_returnsEmptyError() {
        String msg = validator.getErrMessage(null);
        assertNotNull(msg);
        assertTrue(msg.contains("empty"));
    }

    @Test
    void getErrMessage_emptyValue_returnsEmptyError() {
        String msg = validator.getErrMessage("");
        assertNotNull(msg);
        assertTrue(msg.contains("empty"));
    }

    @Test
    void getErrMessage_whitespaceValue_returnsEmptyError() {
        String msg = validator.getErrMessage("   ");
        assertNotNull(msg);
        assertTrue(msg.contains("empty"));
    }

    @Test
    void getErrMessage_nonEmptyValue_returnsTooLongError() {
        String msg = validator.getErrMessage("value");
        assertNotNull(msg);
        assertTrue(msg.contains("too long"));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Lookup Value", validator.getFetchFieldName());
    }

    @Test
    void getFieldType_returnsTextFieldLogicalName() {
        assertEquals("com.krista.fields.Text", validator.getFieldType());
    }

    @Test
    void getFetchStepMessage_returnsNonEmpty() {
        String msg = validator.getFetchStepMessage();
        assertNotNull(msg);
        assertFalse(msg.isEmpty());
    }
}
