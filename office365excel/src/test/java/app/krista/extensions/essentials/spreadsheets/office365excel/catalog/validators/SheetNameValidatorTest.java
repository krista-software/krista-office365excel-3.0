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

import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import com.microsoft.graph.http.GraphServiceException;
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
class SheetNameValidatorTest {

    @Mock
    private WorksheetService worksheetService;

    private SheetNameValidator validator;
    private Map<Validator.ValidationResource, String> context;

    @BeforeEach
    void setUp() {
        validator = new SheetNameValidator(worksheetService);
        context = new HashMap<>();
        context.put(Validator.ValidationResource.FILE_ID, "file1");
        context.put(Validator.ValidationResource.DRIVE_ID, "drive1");
    }

    // --- validate tests ---

    @Test
    void validate_validSheetName_returnsTrue() {
        when(worksheetService.worksheetExists("file1", "Sheet1", "drive1")).thenReturn(true);
        assertTrue(validator.validate("Sheet1", context));
    }

    @Test
    void validate_sheetDoesNotExist_returnsFalse() {
        when(worksheetService.worksheetExists("file1", "Missing", "drive1")).thenReturn(false);
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
    void validate_invalidCharsOpenBracket_returnsFalse() {
        assertFalse(validator.validate("Sheet[1]", context));
    }

    @Test
    void validate_invalidCharsCloseBracket_returnsFalse() {
        assertFalse(validator.validate("Sheet]1", context));
    }

    @Test
    void validate_invalidCharsAsterisk_returnsFalse() {
        assertFalse(validator.validate("Sheet*1", context));
    }

    @Test
    void validate_invalidCharsQuestionMark_returnsFalse() {
        assertFalse(validator.validate("Sheet?1", context));
    }

    @Test
    void validate_invalidCharsForwardSlash_returnsFalse() {
        assertFalse(validator.validate("Sheet/1", context));
    }

    @Test
    void validate_invalidCharsBackslash_returnsFalse() {
        assertFalse(validator.validate("Sheet\\1", context));
    }

    @Test
    void validate_startsWithSingleQuote_returnsFalse() {
        assertFalse(validator.validate("'Sheet1", context));
    }

    @Test
    void validate_endsWithSingleQuote_returnsFalse() {
        assertFalse(validator.validate("Sheet1'", context));
    }

    @Test
    void validate_missingFileIdInContext_returnsFalse() {
        context.remove(Validator.ValidationResource.FILE_ID);
        assertFalse(validator.validate("Sheet1", context));
    }

    @Test
    void validate_emptyFileIdInContext_returnsFalse() {
        context.put(Validator.ValidationResource.FILE_ID, "  ");
        assertFalse(validator.validate("Sheet1", context));
    }

    @Test
    void validate_graphServiceException_rethrown() {
        GraphServiceException gse = mock(GraphServiceException.class);
        when(worksheetService.worksheetExists("file1", "Sheet1", "drive1")).thenThrow(gse);
        assertThrows(GraphServiceException.class, () -> validator.validate("Sheet1", context));
    }

    @Test
    void validate_illegalStateException_rethrown() {
        when(worksheetService.worksheetExists("file1", "Sheet1", "drive1"))
                .thenThrow(new IllegalStateException("Access denied"));
        assertThrows(IllegalStateException.class, () -> validator.validate("Sheet1", context));
    }

    @Test
    void validate_generalException_returnsFalse() {
        when(worksheetService.worksheetExists("file1", "Sheet1", "drive1"))
                .thenThrow(new RuntimeException("unknown error"));
        assertFalse(validator.validate("Sheet1", context));
    }

    @Test
    void validate_nullWorksheetService_skipsExistenceCheck() {
        SheetNameValidator validatorNoService = new SheetNameValidator(null);
        assertTrue(validatorNoService.validate("Sheet1", context));
    }

    // --- interface method tests ---

    @Test
    void getFetchFieldName_returnsExpectedValue() {
        assertEquals("Sheet Name", validator.getFetchFieldName());
    }

    @Test
    void getFieldType_returnsTextFieldLogicalName() {
        assertEquals("com.krista.fields.Text", validator.getFieldType());
    }

    @Test
    void getFetchStepMessage_withService_returnsNonEmpty() {
        String msg = validator.getFetchStepMessage();
        assertNotNull(msg);
        assertFalse(msg.isEmpty());
        assertTrue(msg.contains("existing"));
    }

    @Test
    void getFetchStepMessage_withoutService_returnsNonEmpty() {
        SheetNameValidator validatorNoService = new SheetNameValidator(null);
        String msg = validatorNoService.getFetchStepMessage();
        assertNotNull(msg);
        assertFalse(msg.isEmpty());
        assertTrue(msg.contains("31 characters"));
    }

    @Test
    void getConfirmStepMessage_withService_returnsNonNull() {
        String msg = validator.getConfirmStepMessage("Sheet1", context);
        assertNotNull(msg);
        assertTrue(msg.contains("Sheet1"));
    }

    @Test
    void getConfirmStepMessage_withoutService_returnsNonNull() {
        SheetNameValidator validatorNoService = new SheetNameValidator(null);
        String msg = validatorNoService.getConfirmStepMessage("Sheet1", context);
        assertNotNull(msg);
        assertTrue(msg.contains("Sheet1"));
    }

    @Test
    void getErrMessage_withService_returnsNonNull() {
        String msg = validator.getErrMessage("Sheet1");
        assertNotNull(msg);
        assertTrue(msg.contains("Sheet1"));
    }

    @Test
    void getErrMessage_withoutService_returnsNonNull() {
        SheetNameValidator validatorNoService = new SheetNameValidator(null);
        String msg = validatorNoService.getErrMessage("Sheet1");
        assertNotNull(msg);
        assertTrue(msg.contains("Sheet1"));
    }
}
