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

class CommentValidatorTest {

    private CommentValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new CommentValidator();
        context = new HashMap<>();
    }

    // --- validate tests ---

    @Test
    void validate_validComment_returnsTrue() {
        assertTrue(validator.validate("This is a valid comment", context));
    }

    @Test
    void validate_emptyComment_returnsTrue() {
        // Empty string is valid (not null, length 0 <= 1024)
        assertTrue(validator.validate("", context));
    }

    @Test
    void validate_nullComment_returnsFalse() {
        assertFalse(validator.validate(null, context));
    }

    @Test
    void validate_exactlyMaxLength_returnsTrue() {
        String comment = "a".repeat(1024);
        assertTrue(validator.validate(comment, context));
    }

    @Test
    void validate_exceedsMaxLength_returnsFalse() {
        String comment = "a".repeat(1025);
        assertFalse(validator.validate(comment, context));
    }

    @Test
    void validate_singleCharacter_returnsTrue() {
        assertTrue(validator.validate("x", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Comment", validator.getFetchFieldName());
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

    @Test
    void getConfirmStepMessage_returnsNonNull() {
        assertNotNull(validator.getConfirmStepMessage("comment", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("comment"));
    }

    @Test
    void getErrMessage_containsMaxLength() {
        String msg = validator.getErrMessage("comment");
        assertTrue(msg.contains("1024"));
    }
}
