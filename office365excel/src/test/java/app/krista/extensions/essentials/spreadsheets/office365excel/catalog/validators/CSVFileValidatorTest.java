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

class CSVFileValidatorTest {

    private CSVFileValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new CSVFileValidator();
        context = new HashMap<>();
    }

    // --- validate tests ---

    @Test
    void validate_validCsvJsonArray_returnsTrue() {
        String json = "[{\"name\": \"data.csv\"}]";
        assertTrue(validator.validate(json, context));
    }

    @Test
    void validate_validCsvWithFileNameKey_returnsTrue() {
        String json = "[{\"fileName\": \"report.csv\"}]";
        assertTrue(validator.validate(json, context));
    }

    @Test
    void validate_validCsvWithOriginalNameKey_returnsTrue() {
        String json = "[{\"originalName\": \"export.csv\"}]";
        assertTrue(validator.validate(json, context));
    }

    @Test
    void validate_multipleValidFiles_returnsTrue() {
        String json = "[{\"name\": \"a.csv\"}, {\"name\": \"b.csv\"}]";
        assertTrue(validator.validate(json, context));
    }

    @Test
    void validate_invalidExtension_returnsFalse() {
        String json = "[{\"name\": \"data.xlsx\"}]";
        assertFalse(validator.validate(json, context));
    }

    @Test
    void validate_noExtension_returnsFalse() {
        String json = "[{\"name\": \"noext\"}]";
        assertFalse(validator.validate(json, context));
    }

    @Test
    void validate_dotAtEnd_returnsFalse() {
        String json = "[{\"name\": \"file.\"}]";
        assertFalse(validator.validate(json, context));
    }

    @Test
    void validate_nullInput_returnsFalse() {
        assertFalse(validator.validate(null, context));
    }

    @Test
    void validate_emptyString_returnsFalse() {
        assertFalse(validator.validate("", context));
    }

    @Test
    void validate_whitespaceOnly_returnsFalse() {
        assertFalse(validator.validate("   ", context));
    }

    @Test
    void validate_emptyJsonArray_returnsFalse() {
        assertFalse(validator.validate("[]", context));
    }

    @Test
    void validate_missingNameField_returnsFalse() {
        String json = "[{\"size\": 123}]";
        assertFalse(validator.validate(json, context));
    }

    @Test
    void validate_nullFileNameValue_returnsFalse() {
        String json = "[{\"name\": null}]";
        assertFalse(validator.validate(json, context));
    }

    @Test
    void validate_emptyFileNameValue_returnsFalse() {
        String json = "[{\"name\": \"\"}]";
        assertFalse(validator.validate(json, context));
    }

    @Test
    void validate_invalidJson_returnsFalse() {
        assertFalse(validator.validate("not json", context));
    }

    @Test
    void validate_nullFileInArray_skipsNull() {
        // A null file entry is skipped (continue), but a valid one after it should pass
        String json = "[null, {\"name\": \"data.csv\"}]";
        assertTrue(validator.validate(json, context));
    }

    @Test
    void validate_mixedValidAndInvalid_returnsFalse() {
        String json = "[{\"name\": \"good.csv\"}, {\"name\": \"bad.txt\"}]";
        assertFalse(validator.validate(json, context));
    }

    @Test
    void validate_uppercaseExtension_returnsTrue() {
        String json = "[{\"name\": \"DATA.CSV\"}]";
        // The validator lowercases the filename before checking
        assertTrue(validator.validate(json, context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("CSV File", validator.getFetchFieldName());
    }

    @Test
    void getFieldType_returnsFileFieldLogicalName() {
        assertEquals("com.krista.fields.File", validator.getFieldType());
    }

    @Test
    void getFetchStepMessage_returnsNonEmpty() {
        String msg = validator.getFetchStepMessage();
        assertNotNull(msg);
        assertFalse(msg.isEmpty());
    }

    @Test
    void getConfirmStepMessage_returnsNonNull() {
        assertNotNull(validator.getConfirmStepMessage("test", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("test"));
    }
}
