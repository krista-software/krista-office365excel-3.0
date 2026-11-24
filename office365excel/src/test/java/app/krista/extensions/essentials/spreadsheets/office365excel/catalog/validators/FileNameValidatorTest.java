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

import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.MockedStatic;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class FileNameValidatorTest {

    private FileNameValidator validator;
    private Map<Validator.ValidationResource, String> context;
    private MockedStatic<StringUtil> stringUtilMock;

    @BeforeEach
    void setUp() {
        validator = new FileNameValidator();
        context = new HashMap<>();
        stringUtilMock = mockStatic(StringUtil.class);
    }

    @AfterEach
    void tearDown() {
        stringUtilMock.close();
    }

    // --- validate tests ---

    @Test
    void validate_validExcelFile_returnsTrue() {
        stringUtilMock.when(() -> StringUtil.isExcelFile("report.xlsx")).thenReturn(true);
        assertTrue(validator.validate("report.xlsx", context));
    }

    @Test
    void validate_notExcelFile_returnsFalse() {
        stringUtilMock.when(() -> StringUtil.isExcelFile("data.csv")).thenReturn(false);
        assertFalse(validator.validate("data.csv", context));
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
        assertEquals("Shared File Name", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("file.txt", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("file.txt"));
    }

    @Test
    void getConfirmStepMessage_containsFileName() {
        String msg = validator.getConfirmStepMessage("test.doc", context);
        assertTrue(msg.contains("test.doc"));
    }

    @Test
    void getErrMessage_containsFileName() {
        String msg = validator.getErrMessage("test.doc");
        assertTrue(msg.contains("test.doc"));
    }
}
