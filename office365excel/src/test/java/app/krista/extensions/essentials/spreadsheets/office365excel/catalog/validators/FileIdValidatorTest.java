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
class FileIdValidatorTest {

    @Mock
    private WorkbookService workbookService;

    private FileIdValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new FileIdValidator(workbookService);
        context = new HashMap<>();
        context.put(Validator.ValidationResource.DRIVE_ID, "drive1");
    }

    // --- validate tests ---

    @Test
    void validate_accessibleFile_returnsTrue() {
        when(workbookService.isFileAccessible("drive1", "file123")).thenReturn(true);
        assertTrue(validator.validate("file123", context));
    }

    @Test
    void validate_inaccessibleFile_returnsFalse() {
        when(workbookService.isFileAccessible("drive1", "bad-file")).thenReturn(false);
        assertFalse(validator.validate("bad-file", context));
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
    void validate_serviceThrowsException_returnsFalse() {
        when(workbookService.isFileAccessible(anyString(), anyString())).thenThrow(new RuntimeException("error"));
        assertFalse(validator.validate("file123", context));
    }

    @Test
    void validate_noDriveIdInContext_usesNullDriveId() {
        context.clear();
        when(workbookService.isFileAccessible(null, "file123")).thenReturn(true);
        assertTrue(validator.validate("file123", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("File Id", validator.getFetchFieldName());
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
        assertNotNull(validator.getConfirmStepMessage("file1", context));
    }

    @Test
    void getErrMessage_returnsNonNull() {
        assertNotNull(validator.getErrMessage("file1"));
    }

    @Test
    void getConfirmStepMessage_containsFileId() {
        String msg = validator.getConfirmStepMessage("abc123", context);
        assertTrue(msg.contains("abc123"));
    }

    @Test
    void getErrMessage_containsFileId() {
        String msg = validator.getErrMessage("abc123");
        assertTrue(msg.contains("abc123"));
    }
}
