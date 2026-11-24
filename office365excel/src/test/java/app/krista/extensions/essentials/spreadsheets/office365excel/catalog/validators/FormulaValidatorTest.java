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

class FormulaValidatorTest {

    private FormulaValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new FormulaValidator();
        context = new HashMap<>();
    }

    // --- validate tests ---

    @Test
    void validate_validFormulaStartingWithEquals_returnsTrue() {
        assertTrue(validator.validate("=SUM(A1:A10)", context));
    }

    @Test
    void validate_simpleEqualsFormula_returnsTrue() {
        assertTrue(validator.validate("=A1+B1", context));
    }

    @Test
    void validate_formulaNotStartingWithEquals_returnsFalse() {
        assertFalse(validator.validate("SUM(A1:A10)", context));
    }

    @Test
    void validate_justEqualsSign_returnsTrue() {
        assertTrue(validator.validate("=", context));
    }

    @Test
    void validate_tooLongFormula_returnsFalse() {
        String longFormula = "=" + "A".repeat(8192);
        assertFalse(validator.validate(longFormula, context));
    }

    @Test
    void validate_exactlyMaxLength_returnsTrue() {
        String formula = "=" + "A".repeat(8191);
        assertTrue(validator.validate(formula, context));
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

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Formula", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("=A1", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("=A1"));
    }

    @Test
    void getConfirmStepMessage_containsFormula() {
        String msg = validator.getConfirmStepMessage("=INVALID()", context);
        assertTrue(msg.contains("=INVALID()"));
    }

    @Test
    void getErrMessage_containsFormula() {
        String msg = validator.getErrMessage("=INVALID()");
        assertTrue(msg.contains("=INVALID()"));
    }
}
