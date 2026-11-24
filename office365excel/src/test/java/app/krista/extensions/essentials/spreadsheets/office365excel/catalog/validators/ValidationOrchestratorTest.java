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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Unit tests for {@link ValidationOrchestrator#validate} method.
 * Uses reflection to inject mock validators for isolated testing.
 */
@ExtendWith(MockitoExtension.class)
@DisplayName("ValidationOrchestrator - validate()")
class ValidationOrchestratorTest {

    @Mock private WorkbookService mockWorkbookService;
    @Mock private WorksheetService mockWorksheetService;
    @Mock private Validator mockValidator;

    private ValidationOrchestrator orchestrator;
    @SuppressWarnings("unchecked")
    private Map<Validator.ValidationResource, Validator> validatorsMap;

    @BeforeEach
    @SuppressWarnings("unchecked")
    void setUp() throws Exception {
        orchestrator = new ValidationOrchestrator(mockWorkbookService, mockWorksheetService);

        // Replace internal validators map with a controlled one via reflection
        Field validatorsField = ValidationOrchestrator.class.getDeclaredField("validators");
        validatorsField.setAccessible(true);
        validatorsMap = (Map<Validator.ValidationResource, Validator>) validatorsField.get(orchestrator);
        validatorsMap.clear();
    }

    // ===== EMPTY MAP =====

    @Test
    @DisplayName("Empty resources map returns empty results")
    void emptyResources_returnsEmptyList() {
        List<ValidationOrchestrator.ValidationResult> results = orchestrator.validate(Map.of());
        assertTrue(results.isEmpty());
    }

    // ===== NO MATCHING VALIDATOR (null branch) =====

    @Test
    @DisplayName("Resource with no matching validator is skipped")
    void noMatchingValidator_skipped() {
        // validatorsMap is empty, so no validator exists for FILE_ID
        Map<Validator.ValidationResource, String> resources = new HashMap<>();
        resources.put(Validator.ValidationResource.FILE_ID, "some-file-id");

        List<ValidationOrchestrator.ValidationResult> results = orchestrator.validate(resources);

        assertTrue(results.isEmpty());
    }

    // ===== VALIDATION PASSES =====

    @Nested
    @DisplayName("When validation passes")
    class ValidationPasses {

        @Test
        @DisplayName("Validator returns true - no result added")
        void validatorReturnsTrue_noResultAdded() {
            validatorsMap.put(Validator.ValidationResource.FILE_ID, mockValidator);
            when(mockValidator.validate(eq("valid-file"), anyMap())).thenReturn(true);

            Map<Validator.ValidationResource, String> resources = new HashMap<>();
            resources.put(Validator.ValidationResource.FILE_ID, "valid-file");

            List<ValidationOrchestrator.ValidationResult> results = orchestrator.validate(resources);

            assertTrue(results.isEmpty());
            verify(mockValidator).validate("valid-file", resources);
        }
    }

    // ===== VALIDATION FAILS =====

    @Nested
    @DisplayName("When validation fails")
    class ValidationFails {

        @Test
        @DisplayName("Validator returns false - ValidationResult added with correct fields")
        void validatorReturnsFalse_resultAdded() {
            validatorsMap.put(Validator.ValidationResource.SHEET_NAME, mockValidator);

            when(mockValidator.validate(eq("BadSheet"), anyMap())).thenReturn(false);
            when(mockValidator.getConfirmStepMessage(eq("BadSheet"), anyMap())).thenReturn("Confirm message");
            when(mockValidator.getFetchFieldName()).thenReturn("Sheet Name");
            when(mockValidator.getFetchStepMessage()).thenReturn("Enter sheet name");
            when(mockValidator.getErrMessage("BadSheet")).thenReturn("Sheet not found");
            when(mockValidator.getFieldType()).thenReturn("TEXT");

            Map<Validator.ValidationResource, String> resources = new HashMap<>();
            resources.put(Validator.ValidationResource.SHEET_NAME, "BadSheet");

            List<ValidationOrchestrator.ValidationResult> results = orchestrator.validate(resources);

            assertEquals(1, results.size());
            ValidationOrchestrator.ValidationResult result = results.get(0);
            assertEquals("Confirm message", result.getConfirmStepMessage());
            assertEquals("Sheet Name", result.getFetchFieldName());
            assertEquals("Enter sheet name", result.getFetchStepMessage());
            assertEquals("Sheet not found", result.getErrMessage());
            assertEquals("TEXT", result.getFieldType());
        }
    }

    // ===== ILLEGAL STATE EXCEPTION =====

    @Nested
    @DisplayName("When validator throws IllegalStateException")
    class IllegalStateExceptionPath {

        @Test
        @DisplayName("Should propagate IllegalStateException to caller")
        void illegalStateException_propagated() {
            validatorsMap.put(Validator.ValidationResource.FILE_ID, mockValidator);

            IllegalStateException expected = new IllegalStateException("Access denied (403)");
            when(mockValidator.validate(eq("file-1"), anyMap())).thenThrow(expected);

            Map<Validator.ValidationResource, String> resources = new HashMap<>();
            resources.put(Validator.ValidationResource.FILE_ID, "file-1");

            IllegalStateException thrown = assertThrows(IllegalStateException.class,
                    () -> orchestrator.validate(resources));

            assertSame(expected, thrown);
        }
    }

    // ===== GENERIC EXCEPTION =====

    @Nested
    @DisplayName("When validator throws generic Exception")
    class GenericExceptionPath {

        @Test
        @DisplayName("Should catch exception and add error ValidationResult")
        void genericException_caughtAndResultAdded() {
            validatorsMap.put(Validator.ValidationResource.DRIVE_ID, mockValidator);

            when(mockValidator.validate(eq("drive-1"), anyMap()))
                    .thenThrow(new RuntimeException("Unexpected error"));
            when(mockValidator.getFetchFieldName()).thenReturn("Drive Id");
            when(mockValidator.getFieldType()).thenReturn("TEXT");

            Map<Validator.ValidationResource, String> resources = new HashMap<>();
            resources.put(Validator.ValidationResource.DRIVE_ID, "drive-1");

            List<ValidationOrchestrator.ValidationResult> results = orchestrator.validate(resources);

            assertEquals(1, results.size());
            ValidationOrchestrator.ValidationResult result = results.get(0);
            assertEquals("An error occurred during validation. Would you like to try again?",
                    result.getConfirmStepMessage());
            assertEquals("Drive Id", result.getFetchFieldName());
            assertEquals("Please provide a valid value.", result.getFetchStepMessage());
            assertEquals("Validation error: Unexpected error", result.getErrMessage());
            assertEquals("TEXT", result.getFieldType());
        }
    }

    // ===== MULTIPLE RESOURCES =====

    @Nested
    @DisplayName("Multiple resources in a single call")
    class MultipleResources {

        @Test
        @DisplayName("Mix of pass, fail, and no-validator entries")
        void mixedResults() {
            Validator passingValidator = mock(Validator.class);
            Validator failingValidator = mock(Validator.class);

            validatorsMap.put(Validator.ValidationResource.FILE_ID, passingValidator);
            validatorsMap.put(Validator.ValidationResource.SHEET_NAME, failingValidator);
            // DRIVE_ID has no validator → skipped

            when(passingValidator.validate(eq("good-file"), anyMap())).thenReturn(true);
            when(failingValidator.validate(eq("bad-sheet"), anyMap())).thenReturn(false);
            when(failingValidator.getConfirmStepMessage(eq("bad-sheet"), anyMap())).thenReturn("confirm");
            when(failingValidator.getFetchFieldName()).thenReturn("Sheet Name");
            when(failingValidator.getFetchStepMessage()).thenReturn("fetch");
            when(failingValidator.getErrMessage("bad-sheet")).thenReturn("not found");
            when(failingValidator.getFieldType()).thenReturn("TEXT");

            // Use LinkedHashMap to control iteration order
            Map<Validator.ValidationResource, String> resources = new LinkedHashMap<>();
            resources.put(Validator.ValidationResource.FILE_ID, "good-file");
            resources.put(Validator.ValidationResource.SHEET_NAME, "bad-sheet");
            resources.put(Validator.ValidationResource.DRIVE_ID, "no-validator-for-this");

            List<ValidationOrchestrator.ValidationResult> results = orchestrator.validate(resources);

            // Only the failing validator produces a result; passing and missing are not added
            assertEquals(1, results.size());
            assertEquals("Sheet Name", results.get(0).getFetchFieldName());
        }
    }
}
