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

class ColumnRangeValidatorTest {

    private ColumnRangeValidator validator;
    private Map<Validator.ValidationResource, String> context;
    private MockedStatic<Worksheets> worksheetsMock;

    @BeforeEach
    void setUp() {
        validator = new ColumnRangeValidator();
        context = new HashMap<>();
        worksheetsMock = mockStatic(Worksheets.class);
    }

    @AfterEach
    void tearDown() {
        worksheetsMock.close();
    }

    // --- validate tests ---

    @Test
    void validate_validColumnRange_returnsTrue() {
        worksheetsMock.when(() -> Worksheets.isValidColumnRange("A:C")).thenReturn(true);
        assertTrue(validator.validate("A:C", context));
    }

    @Test
    void validate_invalidColumnRange_returnsFalse() {
        worksheetsMock.when(() -> Worksheets.isValidColumnRange("XYZ")).thenReturn(false);
        assertFalse(validator.validate("XYZ", context));
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
        worksheetsMock.when(() -> Worksheets.isValidColumnRange("A:B")).thenThrow(new RuntimeException("error"));
        assertFalse(validator.validate("A:B", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Column Range", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("A:Z", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("A:Z"));
    }

    @Test
    void getConfirmStepMessage_containsRange() {
        String msg = validator.getConfirmStepMessage("B:D", context);
        assertTrue(msg.contains("B:D"));
    }

    @Test
    void getErrMessage_containsRange() {
        String msg = validator.getErrMessage("B:D");
        assertTrue(msg.contains("B:D"));
    }
}
