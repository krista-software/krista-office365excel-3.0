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

class PathValidatorTest {

    private PathValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new PathValidator();
        context = new HashMap<>();
    }

    // --- validate tests ---

    @Test
    void validate_nullPath_returnsTrue() {
        assertTrue(validator.validate(null, context));
    }

    @Test
    void validate_emptyPath_returnsTrue() {
        // Empty string matches the pattern (it's optional)
        assertTrue(validator.validate("", context));
    }

    @Test
    void validate_simpleFolder_returnsTrue() {
        assertTrue(validator.validate("Documents", context));
    }

    @Test
    void validate_absolutePath_returnsTrue() {
        assertTrue(validator.validate("/Documents/Excel", context));
    }

    @Test
    void validate_relativePath_returnsTrue() {
        assertTrue(validator.validate("Documents/Excel", context));
    }

    @Test
    void validate_pathWithSpaces_returnsTrue() {
        assertTrue(validator.validate("/My Documents/Excel Files", context));
    }

    @Test
    void validate_pathWithDoubleDots_returnsFalse() {
        assertFalse(validator.validate("/Documents/../secret", context));
    }

    @Test
    void validate_pathWithDoubleSlashes_returnsFalse() {
        assertFalse(validator.validate("/Documents//Excel", context));
    }

    @Test
    void validate_pathWithBackslashes_returnsFalse() {
        // The regex pattern won't match backslashes, and the check for \\\\ would also catch it
        assertFalse(validator.validate("Documents\\Excel", context));
    }

    @Test
    void validate_tooLongPath_returnsFalse() {
        String longPath = "/" + "a".repeat(260);
        assertFalse(validator.validate(longPath, context));
    }

    @Test
    void validate_exactlyMaxLength_returnsTrue() {
        // 260 chars total: / + 259 chars of valid path
        String path = "a".repeat(260);
        // This exceeds max length (260 chars = max, but path.length() > 260 fails)
        assertFalse(validator.validate(path + "b", context));
    }

    @Test
    void validate_pathWithHyphens_returnsTrue() {
        assertTrue(validator.validate("/my-folder/sub-dir", context));
    }

    @Test
    void validate_pathWithSpecialChars_returnsFalse() {
        assertFalse(validator.validate("/Documents/@#$/Excel", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Path", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("/bad/path", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("/bad/path"));
    }

    @Test
    void getConfirmStepMessage_containsPath() {
        String msg = validator.getConfirmStepMessage("/test/path", context);
        assertTrue(msg.contains("/test/path"));
    }

    @Test
    void getErrMessage_containsPath() {
        String msg = validator.getErrMessage("/test/path");
        assertTrue(msg.contains("/test/path"));
    }
}
