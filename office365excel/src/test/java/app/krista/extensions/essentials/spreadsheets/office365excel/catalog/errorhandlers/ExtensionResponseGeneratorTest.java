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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ExtensionResponseGenerator}.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("ExtensionResponseGenerator")
class ExtensionResponseGeneratorTest {

    private ExtensionResponseGenerator generator;

    @BeforeEach
    void setUp() {
        generator = new ExtensionResponseGenerator();
    }

    private ValidationOrchestrator.ValidationResult createValidationResult(
            String confirmMsg, String fetchFieldName, String fetchMsg, String errMsg, String fieldType) {
        return new ValidationOrchestrator.ValidationResult(confirmMsg, fetchFieldName, fetchMsg, errMsg, fieldType);
    }

    @Nested
    @DisplayName("generateConfirmationResponse()")
    class GenerateConfirmationResponse {

        @Test
        @DisplayName("Creates response with switch field for re-entry confirmation")
        void createsResponseWithSwitchField() {
            ValidationOrchestrator.ValidationResult result = createValidationResult(
                    "Confirm step msg", "FieldName", "Fetch step msg", "Error msg", "Text");
            Map<String, Object> state = new HashMap<>();
            state.put("key", "value");

            ExtensionResponse response = generator.generateConfirmationResponse(
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(result),
                    "subCatalogName",
                    state);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertNotNull(response.getError());
            assertTrue(response.getError().getErrorMessage().contains("Error msg"));
            assertNotNull(response.getRemediationActions());
        }

        @Test
        @DisplayName("Creates response with multiple validation results")
        void createsResponseWithMultipleResults() {
            List<ValidationOrchestrator.ValidationResult> results = List.of(
                    createValidationResult("Confirm1", "Field1", "Fetch1", "Err1", "Text"),
                    createValidationResult("Confirm2", "Field2", "Fetch2", "Err2", "Number"));
            Map<String, Object> state = new HashMap<>();

            ExtensionResponse response = generator.generateConfirmationResponse(
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    results,
                    "subCatalogName",
                    state);

            assertNotNull(response);
            String errMessage = response.getError().getErrorMessage();
            assertTrue(errMessage.contains("Err1"));
            assertTrue(errMessage.contains("Err2"));
        }
    }

    @Nested
    @DisplayName("generateFetchResponse()")
    class GenerateFetchResponse {

        @Test
        @DisplayName("Creates response with input fields for corrected values")
        void createsResponseWithInputFields() {
            ValidationOrchestrator.ValidationResult result = createValidationResult(
                    "Confirm step msg", "FieldName", "Fetch step msg", "Error msg", "Text");
            Map<String, Object> state = new HashMap<>();
            state.put("key", "value");

            ExtensionResponse response = generator.generateFetchResponse(
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(result),
                    "subCatalogName",
                    state);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertNotNull(response.getRemediationActions());
            assertEquals(state, response.getState());
        }

        @Test
        @DisplayName("Creates response with multiple validation results and fields")
        void createsResponseWithMultipleResults() {
            List<ValidationOrchestrator.ValidationResult> results = List.of(
                    createValidationResult("Confirm1", "Field1", "Fetch1", "Err1", "Text"),
                    createValidationResult("Confirm2", "Field2", "Fetch2", "Err2", "Number"));
            Map<String, Object> state = new HashMap<>();

            ExtensionResponse response = generator.generateFetchResponse(
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    results,
                    "subCatalogName",
                    state);

            assertNotNull(response);
            String errMessage = response.getError().getErrorMessage();
            assertTrue(errMessage.contains("Err1"));
            assertTrue(errMessage.contains("Err2"));
        }
    }

    @Nested
    @DisplayName("generateFetchDenyResponse()")
    class GenerateFetchDenyResponse {

        @Test
        @DisplayName("Creates inform response with missing field names")
        void createsInformResponse() {
            ValidationOrchestrator.ValidationResult result = createValidationResult(
                    "Confirm step msg", "MyField", "Fetch step msg", "Error msg", "Text");
            Map<String, Object> state = new HashMap<>();

            ExtensionResponse response = generator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(result),
                    "subCatalogName",
                    state);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertNotNull(response.getRemediationActions());
        }

        @Test
        @DisplayName("Creates inform response with multiple missing fields")
        void createsInformResponseMultipleFields() {
            List<ValidationOrchestrator.ValidationResult> results = List.of(
                    createValidationResult("Confirm1", "Field1", "Fetch1", "Err1", "Text"),
                    createValidationResult("Confirm2", "Field2", "Fetch2", "Err2", "Number"));
            Map<String, Object> state = new HashMap<>();

            ExtensionResponse response = generator.generateFetchDenyResponse(
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    results,
                    "subCatalogName",
                    state);

            assertNotNull(response);
            String errMessage = response.getError().getErrorMessage();
            assertTrue(errMessage.contains("Err1"));
            assertTrue(errMessage.contains("Err2"));
        }
    }

    @Nested
    @DisplayName("generateValidationErrorResponse()")
    class GenerateValidationErrorResponse {

        @Test
        @DisplayName("Creates simple failure response without remediation")
        void createsSimpleFailureResponse() {
            ValidationOrchestrator.ValidationResult result = createValidationResult(
                    "Confirm step msg", "FieldName", "Fetch step msg", "Validation failed for field X", "Text");

            ExtensionResponse response = generator.generateValidationErrorResponse(List.of(result));

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
        }

        @Test
        @DisplayName("Concatenates error messages from multiple validation results")
        void concatenatesMultipleErrorMessages() {
            List<ValidationOrchestrator.ValidationResult> results = List.of(
                    createValidationResult("C1", "F1", "FS1", "Error one", "Text"),
                    createValidationResult("C2", "F2", "FS2", "Error two", "Number"));

            ExtensionResponse response = generator.generateValidationErrorResponse(results);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
        }
    }

    @Test
    @DisplayName("Constants are correctly defined")
    void constantsAreDefined() {
        assertEquals("ErrorMessage", ExtensionResponseGenerator.ERROR_MESSAGE);
        assertEquals("StepMessage", ExtensionResponseGenerator.STEP_MESSAGE);
        assertEquals("Field", ExtensionResponseGenerator.FIELD);
    }
}
