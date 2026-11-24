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

class FormulaNameValidatorTest {

    private FormulaNameValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new FormulaNameValidator();
        context = new HashMap<>();
    }

    // --- validate tests ---

    @Test
    void validate_validSimpleName_returnsTrue() {
        assertTrue(validator.validate("MyFormula", context));
    }

    @Test
    void validate_nameStartingWithUnderscore_returnsTrue() {
        assertTrue(validator.validate("_myFormula", context));
    }

    @Test
    void validate_nameWithNumbersAndDots_returnsTrue() {
        assertTrue(validator.validate("Formula1.v2", context));
    }

    @Test
    void validate_nameWithUnderscores_returnsTrue() {
        assertTrue(validator.validate("my_formula_name", context));
    }

    @Test
    void validate_nameStartingWithNumber_returnsFalse() {
        assertFalse(validator.validate("1Formula", context));
    }

    @Test
    void validate_nameWithSpecialChars_returnsFalse() {
        assertFalse(validator.validate("Formula@Name", context));
    }

    @Test
    void validate_nameWithSpaces_returnsFalse() {
        assertFalse(validator.validate("My Formula", context));
    }

    @Test
    void validate_nameWithHyphen_returnsFalse() {
        assertFalse(validator.validate("my-formula", context));
    }

    @Test
    void validate_tooLongName_returnsFalse() {
        String longName = "a".repeat(256);
        assertFalse(validator.validate(longName, context));
    }

    @Test
    void validate_exactlyMaxLength_returnsTrue() {
        String maxName = "a".repeat(255);
        assertTrue(validator.validate(maxName, context));
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
    void validate_singleLetter_returnsTrue() {
        assertTrue(validator.validate("A", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Formula Name", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("test", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("test"));
    }

    @Test
    void getConfirmStepMessage_containsFormulaName() {
        String msg = validator.getConfirmStepMessage("Bad@Name", context);
        assertTrue(msg.contains("Bad@Name"));
    }

    @Test
    void getErrMessage_containsFormulaName() {
        String msg = validator.getErrMessage("Bad@Name");
        assertTrue(msg.contains("Bad@Name"));
    }
}
