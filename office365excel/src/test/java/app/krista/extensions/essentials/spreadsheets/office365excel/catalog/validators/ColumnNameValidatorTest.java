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

import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.quality.Strictness;
import org.mockito.junit.jupiter.MockitoSettings;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class ColumnNameValidatorTest {

    @Mock
    private WorkbookService workbookService;

    private ColumnNameValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new ColumnNameValidator(workbookService);
        context = new HashMap<>();
        context.put(Validator.ValidationResource.DRIVE_ID, "drive1");
        context.put(Validator.ValidationResource.FILE_ID, "file1");
        context.put(Validator.ValidationResource.TABLE_NAME, "Table1");
    }

    // --- validate tests ---

    @Test
    void validate_validColumn_returnsTrue() {
        when(workbookService.isColumnExistsInTable("drive1", "file1", "Table1", "Name")).thenReturn(true);
        assertTrue(validator.validate("Name", context));
    }

    @Test
    void validate_columnNotInTable_returnsFalse() {
        when(workbookService.isColumnExistsInTable("drive1", "file1", "Table1", "Missing")).thenReturn(false);
        assertFalse(validator.validate("Missing", context));
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
    void validate_missingDriveIdInContext_returnsFalse() {
        context.remove(Validator.ValidationResource.DRIVE_ID);
        assertFalse(validator.validate("Name", context));
    }

    @Test
    void validate_missingFileIdInContext_returnsFalse() {
        context.remove(Validator.ValidationResource.FILE_ID);
        assertFalse(validator.validate("Name", context));
    }

    @Test
    void validate_missingTableNameInContext_returnsFalse() {
        context.remove(Validator.ValidationResource.TABLE_NAME);
        assertFalse(validator.validate("Name", context));
    }

    @Test
    void validate_serviceThrowsException_returnsFalse() {
        when(workbookService.isColumnExistsInTable(anyString(), anyString(), anyString(), anyString()))
                .thenThrow(new RuntimeException("API error"));
        assertFalse(validator.validate("Name", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Column Name", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("Col1", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("Col1"));
    }

    @Test
    void getConfirmStepMessage_containsColumnName() {
        String msg = validator.getConfirmStepMessage("Revenue", context);
        assertTrue(msg.contains("Revenue"));
    }

    @Test
    void getErrMessage_containsColumnName() {
        String msg = validator.getErrMessage("Revenue");
        assertTrue(msg.contains("Revenue"));
    }
}
