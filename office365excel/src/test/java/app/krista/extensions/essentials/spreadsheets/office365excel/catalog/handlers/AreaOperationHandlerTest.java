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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.ksdk.context.RequestContext;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.*;
import java.util.function.Function;
import java.util.function.Supplier;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Comprehensive unit tests for {@link AreaOperationHandler#execute} method.
 * Covers all reachable code paths for 100% code coverage.
 */
@ExtendWith(MockitoExtension.class)
@DisplayName("AreaOperationHandler Tests")
class AreaOperationHandlerTest {

    @Mock
    private ValidationOrchestrator mockValidationOrchestrator;

    @Mock
    private ErrorHandlingStateManager mockStateManager;

    @Mock
    private ExtensionResponseGenerator mockResponseGenerator;

    @Mock
    private RequestContext mockRequestContext;

    @Mock
    private TelemetryHelper mockTelemetryHelper;

    @Mock
    private Supplier<String> mockOperation;

    @Mock
    private Function<String, Map<String, Object>> mockResultMapper;

    @Mock
    private Function<String, Map<String, Object>> mockStateMapper;

    private AreaOperationHandler<String> handler;

    // Common test data
    private static final String OPERATION_NAME = "testOperation";
    private static final String CONFIRMATION_ACTION = SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET;

    @BeforeEach
    void setUp() {
        handler = new AreaOperationHandler<>(
                mockValidationOrchestrator,
                mockStateManager,
                mockResponseGenerator,
                mockRequestContext,
                mockTelemetryHelper
        );
    }

    private Map<Validator.ValidationResource, String> createValidationMap() {
        Map<Validator.ValidationResource, String> map = new HashMap<>();
        map.put(Validator.ValidationResource.FILE_ID, "test-file-id");
        map.put(Validator.ValidationResource.SHEET_NAME, "Sheet1");
        return map;
    }

    private Map<String, Object> createStateValues() {
        Map<String, Object> state = new HashMap<>();
        state.put("fileId", "test-file-id");
        state.put("sheetName", "Sheet1");
        return state;
    }

    private Map<String, String> createTelemetryTags() {
        Map<String, String> tags = new HashMap<>();
        tags.put("customTag", "customValue");
        return tags;
    }

    private List<ValidationOrchestrator.ValidationResult> createValidationResults() {
        return List.of(
                new ValidationOrchestrator.ValidationResult(
                        "confirmMsg", "File Id", "fetchMsg", "File not found", "TEXT"
                )
        );
    }

    private List<ValidationOrchestrator.ValidationResult> createMultipleValidationResults() {
        return List.of(
                new ValidationOrchestrator.ValidationResult(
                        "confirmMsg1", "File Id", "fetchMsg1", "File not found", "TEXT"
                ),
                new ValidationOrchestrator.ValidationResult(
                        "confirmMsg2", "Sheet Name", "fetchMsg2", "Sheet not found", "TEXT"
                )
        );
    }

    // ===== PATH 1: allowRetry=FALSE/NULL → skip validation → operation succeeds =====

    @Nested
    @DisplayName("When allowRetry is disabled (FALSE/NULL)")
    class AllowRetryDisabled {

        @Test
        @DisplayName("allowRetry=FALSE - skips validation and executes operation successfully")
        void allowRetryFalse_skipsValidation_executesOperation() {
            when(mockOperation.get()).thenReturn("success");
            when(mockResultMapper.apply("success")).thenReturn(Map.of("result", "success"));

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            assertNotNull(response);
            verify(mockValidationOrchestrator, never()).validate(anyMap());
            verify(mockOperation).get();
            verify(mockResultMapper).apply("success");
            verify(mockTelemetryHelper).recordSuccess(eq("office365Excel." + OPERATION_NAME), anyLong(), anyMap());
        }

        @Test
        @DisplayName("allowRetry=NULL - skips validation and executes operation successfully")
        void allowRetryNull_skipsValidation_executesOperation() {
            when(mockOperation.get()).thenReturn("result");
            when(mockResultMapper.apply("result")).thenReturn(Map.of("data", "value"));

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), null
            );

            assertNotNull(response);
            verify(mockValidationOrchestrator, never()).validate(anyMap());
            verify(mockOperation).get();
        }
    }

    // ===== PATH 2: allowRetry=TRUE → validation passes → operation succeeds =====

    @Nested
    @DisplayName("When allowRetry=TRUE and validation passes")
    class AllowRetryEnabledValidationPasses {

        @Test
        @DisplayName("Should execute operation and return success response")
        void validationPasses_executesOperation() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            when(mockValidationOrchestrator.validate(validationMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenReturn("operationResult");
            when(mockResultMapper.apply("operationResult")).thenReturn(Map.of("key", "value"));

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            assertNotNull(response);
            verify(mockValidationOrchestrator).validate(validationMap);
            verify(mockOperation).get();
            verify(mockResultMapper).apply("operationResult");
            verify(mockTelemetryHelper).recordSuccess(eq("office365Excel." + OPERATION_NAME), anyLong(), anyMap());
            verify(mockStateManager, never()).put(anyString(), anyMap());
            verify(mockResponseGenerator, never()).generateConfirmationResponse(any(), anyList(), anyString(), anyMap());
        }
    }

    // ===== PATH 3: allowRetry=TRUE → validation fails → confirmation response =====

    @Nested
    @DisplayName("When allowRetry=TRUE and validation fails")
    class AllowRetryEnabledValidationFails {

        @Test
        @DisplayName("Should store state and return confirmation response for single validation error")
        void validationFails_returnsConfirmationResponse() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            List<ValidationOrchestrator.ValidationResult> validationResults = createValidationResults();
            ExtensionResponse expectedResponse = mock(ExtensionResponse.class);

            when(mockValidationOrchestrator.validate(validationMap)).thenReturn(validationResults);
            when(mockStateMapper.apply(anyString())).thenReturn(Map.of("stateId", "id"));
            when(mockResponseGenerator.generateConfirmationResponse(
                    any(ExtensionResponse.Error.ExceptionType.class),
                    anyList(), eq(CONFIRMATION_ACTION), anyMap()
            )).thenReturn(expectedResponse);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            assertEquals(expectedResponse, response);
            verify(mockStateManager).put(anyString(), anyMap());
            verify(mockOperation, never()).get();
            verify(mockTelemetryHelper).recordValidationError(
                    eq("office365Excel." + OPERATION_NAME), anyLong(), anyString(), anyMap()
            );
        }

        @Test
        @DisplayName("Should handle multiple validation errors and store all in state")
        void multipleValidationErrors_storesAllAndReturnsConfirmation() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            List<ValidationOrchestrator.ValidationResult> validationResults = createMultipleValidationResults();
            ExtensionResponse expectedResponse = mock(ExtensionResponse.class);

            when(mockValidationOrchestrator.validate(validationMap)).thenReturn(validationResults);
            when(mockStateMapper.apply(anyString())).thenReturn(Map.of("stateId", "id"));
            when(mockResponseGenerator.generateConfirmationResponse(
                    any(ExtensionResponse.Error.ExceptionType.class),
                    anyList(), eq(CONFIRMATION_ACTION), anyMap()
            )).thenReturn(expectedResponse);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            assertEquals(expectedResponse, response);
            verify(mockStateManager).put(anyString(), argThat(state ->
                    state.containsKey(SubCatalogConstants.VALIDATION_RESULTS)
            ));
        }
    }

    // ===== PATH 4: allowRetry=TRUE → validation throws IllegalStateException =====

    @Nested
    @DisplayName("When validation throws IllegalStateException")
    class ValidationThrowsIllegalStateException {

        @Test
        @DisplayName("Should return error response with access error message")
        void validationThrowsIllegalState_returnsAccessError() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            IllegalStateException exception = new IllegalStateException("Access denied to workbook");

            when(mockValidationOrchestrator.validate(validationMap)).thenThrow(exception);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            assertNotNull(response);
            verify(mockOperation, never()).get();
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OPERATION_NAME), anyLong(), eq(exception), anyMap()
            );
        }

        @Test
        @DisplayName("Should return error response when IllegalStateException has null message")
        void validationThrowsIllegalStateNullMessage_returnsAccessError() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            IllegalStateException exception = new IllegalStateException();

            when(mockValidationOrchestrator.validate(validationMap)).thenThrow(exception);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            assertNotNull(response);
            verify(mockOperation, never()).get();
        }
    }

    // ===== PATH 5: Operation throws MustAuthorizeException =====

    @Nested
    @DisplayName("When operation throws MustAuthorizeException")
    class OperationThrowsMustAuthorizeException {

        @Test
        @DisplayName("invokeAsUser=false - should return authorization error response")
        void mustAuthorize_invokeAsUserFalse_returnsAuthErrorResponse() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            MustAuthorizeException authException = new MustAuthorizeException("Token expired");

            when(mockValidationOrchestrator.validate(validationMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            assertNotNull(response);
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OPERATION_NAME), anyLong(), eq(authException), anyMap()
            );
        }

        @Test
        @DisplayName("invokeAsUser=true - should rethrow MustAuthorizeException")
        void mustAuthorize_invokeAsUserTrue_rethrowsException() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            MustAuthorizeException authException = new MustAuthorizeException("Auth required");

            when(mockValidationOrchestrator.validate(validationMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(true);

            assertThrows(MustAuthorizeException.class, () ->
                    handler.execute(
                            OPERATION_NAME, validationMap, createStateValues(),
                            CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                            createTelemetryTags(), Boolean.TRUE
                    )
            );
        }

        @Test
        @DisplayName("invokeAsUser=false with null message - should handle gracefully")
        void mustAuthorize_nullMessage_returnsAuthErrorResponse() {
            when(mockOperation.get()).thenThrow(new MustAuthorizeException((String) null));
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            assertNotNull(response);
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OPERATION_NAME), anyLong(), any(MustAuthorizeException.class), anyMap()
            );
        }
    }

    // ===== PATH 6: Operation throws generic Exception =====

    @Nested
    @DisplayName("When operation throws generic Exception")
    class OperationThrowsGenericException {

        @Test
        @DisplayName("Should return error response and remove state")
        void genericException_returnsErrorAndRemovesState() {
            RuntimeException exception = new RuntimeException("Graph API error");

            when(mockOperation.get()).thenThrow(exception);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            assertNotNull(response);
            verify(mockStateManager).remove(anyString());
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OPERATION_NAME), anyLong(), eq(exception), anyMap()
            );
        }

        @Test
        @DisplayName("Should handle exception with null message")
        void genericExceptionNullMessage_returnsError() {
            RuntimeException exception = new RuntimeException((String) null);

            when(mockOperation.get()).thenThrow(exception);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            assertNotNull(response);
            verify(mockStateManager).remove(anyString());
        }

        @Test
        @DisplayName("Should handle exception when allowRetry=TRUE and validation passed")
        void genericException_afterValidationPasses_returnsError() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            RuntimeException exception = new RuntimeException("Unexpected error");

            when(mockValidationOrchestrator.validate(validationMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenThrow(exception);

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            assertNotNull(response);
            verify(mockValidationOrchestrator).validate(validationMap);
            verify(mockStateManager).remove(anyString());
        }
    }

    // ===== TELEMETRY & EDGE CASES =====

    @Nested
    @DisplayName("Telemetry and edge cases")
    class TelemetryAndEdgeCases {

        @Test
        @DisplayName("Should handle null telemetryTags gracefully")
        void nullTelemetryTags_shouldNotThrow() {
            when(mockOperation.get()).thenReturn("result");
            when(mockResultMapper.apply("result")).thenReturn(Map.of("key", "val"));

            ExtensionResponse response = handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    null, Boolean.FALSE
            );

            assertNotNull(response);
            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel." + OPERATION_NAME), anyLong(), anyMap()
            );
        }

        @Test
        @DisplayName("Should build correct base metric from operation name")
        void baseMetric_includesOperationName() {
            String customOpName = "getCellData";
            when(mockOperation.get()).thenReturn("data");
            when(mockResultMapper.apply("data")).thenReturn(Map.of("cell", "A1"));

            handler.execute(
                    customOpName, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel.getCellData"), anyLong(), anyMap()
            );
        }

        @Test
        @DisplayName("Should include validation resource count and state value count in telemetry tags")
        void telemetryTags_includeResourceAndStateCounts() {
            when(mockOperation.get()).thenReturn("ok");
            when(mockResultMapper.apply("ok")).thenReturn(Map.of());

            Map<Validator.ValidationResource, String> validationMap = createValidationMap(); // 2 entries
            Map<String, Object> stateValues = createStateValues(); // 2 entries

            handler.execute(
                    OPERATION_NAME, validationMap, stateValues,
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    Map.of(), Boolean.FALSE
            );

            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel." + OPERATION_NAME),
                    anyLong(),
                    argThat(tags ->
                            "2".equals(tags.get("validationResourceCount")) &&
                            "2".equals(tags.get("stateValueCount")) &&
                            OPERATION_NAME.equals(tags.get("operation"))
                    )
            );
        }

        @Test
        @DisplayName("Should preserve custom telemetry tags alongside generated ones")
        void telemetryTags_preservesCustomTags() {
            when(mockOperation.get()).thenReturn("ok");
            when(mockResultMapper.apply("ok")).thenReturn(Map.of());

            Map<String, String> customTags = new HashMap<>();
            customTags.put("myTag", "myValue");

            handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    customTags, Boolean.FALSE
            );

            verify(mockTelemetryHelper).recordSuccess(
                    anyString(), anyLong(),
                    argThat(tags -> "myValue".equals(tags.get("myTag")))
            );
        }

        @Test
        @DisplayName("recordValidationFailure - should record error count and failed fields in telemetry")
        void validationFailure_recordsErrorCountAndFields() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            List<ValidationOrchestrator.ValidationResult> results = createMultipleValidationResults();
            ExtensionResponse expectedResponse = mock(ExtensionResponse.class);

            when(mockValidationOrchestrator.validate(validationMap)).thenReturn(results);
            when(mockStateMapper.apply(anyString())).thenReturn(Map.of());
            when(mockResponseGenerator.generateConfirmationResponse(any(), anyList(), anyString(), anyMap()))
                    .thenReturn(expectedResponse);

            handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            verify(mockTelemetryHelper).recordValidationError(
                    eq("office365Excel." + OPERATION_NAME),
                    anyLong(),
                    argThat(msg -> msg.contains("2 errors")),
                    argThat(tags ->
                            "2".equals(tags.get("validationErrorCount")) &&
                            tags.get("validationErrors") != null
                    )
            );
        }

        @Test
        @DisplayName("recordAuthorizationError - should set errorType AUTHORIZATION and errorClass in tags")
        void authorizationError_setsCorrectTelemetryTags() {
            MustAuthorizeException authException = new MustAuthorizeException("token expired");

            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OPERATION_NAME),
                    anyLong(),
                    eq(authException),
                    argThat(tags ->
                            "AUTHORIZATION".equals(tags.get("errorType")) &&
                            "MustAuthorizeException".equals(tags.get("errorClass")) &&
                            "token expired".equals(tags.get("errorMessage"))
                    )
            );
        }

        @Test
        @DisplayName("recordAuthorizationError - null message should default to 'Authorization required'")
        void authorizationErrorNullMessage_defaultsMessage() {
            MustAuthorizeException authException = new MustAuthorizeException((String) null);

            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            verify(mockTelemetryHelper).recordError(
                    anyString(), anyLong(), any(),
                    argThat(tags -> "Authorization required".equals(tags.get("errorMessage")))
            );
        }

        @Test
        @DisplayName("recordGeneralError - should set errorType LOGIC_ERROR in tags")
        void generalError_setsCorrectTelemetryTags() {
            RuntimeException exception = new RuntimeException("Something went wrong");

            when(mockOperation.get()).thenThrow(exception);

            handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OPERATION_NAME),
                    anyLong(),
                    eq(exception),
                    argThat(tags ->
                            "LOGIC_ERROR".equals(tags.get("errorType")) &&
                            "RuntimeException".equals(tags.get("errorClass")) &&
                            "Something went wrong".equals(tags.get("errorMessage"))
                    )
            );
        }

        @Test
        @DisplayName("recordGeneralError - null message should default to 'Unknown error'")
        void generalErrorNullMessage_defaultsMessage() {
            RuntimeException exception = new RuntimeException((String) null);

            when(mockOperation.get()).thenThrow(exception);

            handler.execute(
                    OPERATION_NAME, createValidationMap(), createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.FALSE
            );

            verify(mockTelemetryHelper).recordError(
                    anyString(), anyLong(), any(),
                    argThat(tags -> "Unknown error".equals(tags.get("errorMessage")))
            );
        }

        @Test
        @DisplayName("recordGeneralError from IllegalStateException during validation")
        void illegalStateValidation_recordsGeneralError() {
            Map<Validator.ValidationResource, String> validationMap = createValidationMap();
            IllegalStateException exception = new IllegalStateException("Infrastructure down");

            when(mockValidationOrchestrator.validate(validationMap)).thenThrow(exception);

            handler.execute(
                    OPERATION_NAME, validationMap, createStateValues(),
                    CONFIRMATION_ACTION, mockStateMapper, mockOperation, mockResultMapper,
                    createTelemetryTags(), Boolean.TRUE
            );

            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OPERATION_NAME),
                    anyLong(),
                    eq(exception),
                    argThat(tags ->
                            "LOGIC_ERROR".equals(tags.get("errorType")) &&
                            "IllegalStateException".equals(tags.get("errorClass"))
                    )
            );
        }
    }
}
