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

import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.MockedStatic;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class RangeValidatorTest {

    private RangeValidator validator;
    private Map<Validator.ValidationResource, String> context;
    private MockedStatic<Worksheets> worksheetsMock;

    @BeforeEach
    void setUp() {
        validator = new RangeValidator();
        context = new HashMap<>();
        worksheetsMock = mockStatic(Worksheets.class);
    }

    @AfterEach
    void tearDown() {
        worksheetsMock.close();
    }

    // --- validate tests ---

    @Test
    void validate_validRange_returnsTrue() {
        worksheetsMock.when(() -> Worksheets.isValidRange("A1:B10")).thenReturn(true);
        assertTrue(validator.validate("A1:B10", context));
    }

    @Test
    void validate_invalidRange_returnsFalse() {
        worksheetsMock.when(() -> Worksheets.isValidRange("INVALID")).thenReturn(false);
        assertFalse(validator.validate("INVALID", context));
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
    void validate_worksheetsThrowsException_returnsFalse() {
        worksheetsMock.when(() -> Worksheets.isValidRange("A1:B10")).thenThrow(new RuntimeException("error"));
        assertFalse(validator.validate("A1:B10", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Range", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("A1:Z99", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("A1:Z99"));
    }

    @Test
    void getConfirmStepMessage_containsRange() {
        String msg = validator.getConfirmStepMessage("C3:D4", context);
        assertTrue(msg.contains("C3:D4"));
    }

    @Test
    void getErrMessage_containsRange() {
        String msg = validator.getErrMessage("C3:D4");
        assertTrue(msg.contains("C3:D4"));
    }
}
