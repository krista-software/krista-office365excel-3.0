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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.handlers.AreaOperationHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
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
 * 100% code coverage tests for {@link AreaOperationHandler#execute} method.
 *
 * Covers every reachable code path:
 * 1. allowRetry=FALSE/NULL → skips validation entirely → executes operation
 * 2. allowRetry=TRUE + validation passes → executes operation
 * 3. allowRetry=TRUE + validation fails → generates confirmation response (3-step retry)
 * 4. allowRetry=TRUE + validation throws IllegalStateException → access error response
 * 5. Operation throws MustAuthorizeException → auth error handling (invokeAsUser true/false)
 * 6. Operation throws generic Exception → error response + state cleanup
 * 7. Telemetry helper methods (prepareTelemetryTags, recordValidationFailure, recordSuccess,
 *    recordAuthorizationError, recordGeneralError) — all tag combinations and null edge cases
 */
@ExtendWith(MockitoExtension.class)
@DisplayName("AllowRetryFeatureTest - AreaOperationHandler.execute() 100% Coverage")
class AllowRetryFeatureTest {

    @Mock private ValidationOrchestrator mockValidationOrchestrator;
    @Mock private ErrorHandlingStateManager mockStateManager;
    @Mock private ExtensionResponseGenerator mockResponseGenerator;
    @Mock private RequestContext mockRequestContext;
    @Mock private TelemetryHelper mockTelemetryHelper;
    @Mock private Supplier<String> mockOperation;
    @Mock private Function<String, Map<String, Object>> mockResultMapper;
    @Mock private Function<String, Map<String, Object>> mockStateMapper;

    private AreaOperationHandler<String> handler;

    private static final String OP_NAME = "testOperation";
    private static final String CONFIRM_ACTION = SubCatalogConstants.CONFIRM_REENTER_DELETE_SHEET;

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

    // ── Helper factories ──────────────────────────────────────────────

    private Map<Validator.ValidationResource, String> validationMap() {
        Map<Validator.ValidationResource, String> map = new HashMap<>();
        map.put(Validator.ValidationResource.FILE_ID, "test-file-id");
        map.put(Validator.ValidationResource.SHEET_NAME, "Sheet1");
        return map;
    }

    private Map<String, Object> stateValues() {
        Map<String, Object> state = new HashMap<>();
        state.put("fileId", "test-file-id");
        state.put("sheetName", "Sheet1");
        return state;
    }

    private Map<String, String> telemetryTags() {
        Map<String, String> tags = new HashMap<>();
        tags.put("customTag", "customValue");
        return tags;
    }

    private List<ValidationOrchestrator.ValidationResult> singleValidationError() {
        return List.of(new ValidationOrchestrator.ValidationResult(
                "confirmMsg", "File Id", "fetchMsg", "File not found", "TEXT"
        ));
    }

    private List<ValidationOrchestrator.ValidationResult> multipleValidationErrors() {
        return List.of(
                new ValidationOrchestrator.ValidationResult(
                        "confirmMsg1", "File Id", "fetchMsg1", "File not found", "TEXT"),
                new ValidationOrchestrator.ValidationResult(
                        "confirmMsg2", "Sheet Name", "fetchMsg2", "Sheet not found", "TEXT")
        );
    }

    private ExtensionResponse execute(Boolean allowRetry) {
        return execute(OP_NAME, validationMap(), stateValues(), telemetryTags(), allowRetry);
    }

    private ExtensionResponse execute(String opName, Map<Validator.ValidationResource, String> vMap,
                                      Map<String, Object> sValues, Map<String, String> tags, Boolean allowRetry) {
        return handler.execute(opName, vMap, sValues, CONFIRM_ACTION,
                mockStateMapper, mockOperation, mockResultMapper, tags, allowRetry);
    }

    // ═══════════════════════════════════════════════════════════════════
    // PATH 1: allowRetry = FALSE / NULL  → skip validation → operation
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("When allowRetry is disabled (FALSE / NULL)")
    class AllowRetryDisabled {

        @Test
        @DisplayName("allowRetry=FALSE → skips validation, executes operation, returns success")
        void allowRetryFalse_skipsValidation_executesOperation() {
            when(mockOperation.get()).thenReturn("success");
            when(mockResultMapper.apply("success")).thenReturn(Map.of("result", "success"));

            ExtensionResponse response = execute(Boolean.FALSE);

            assertNotNull(response);
            verify(mockValidationOrchestrator, never()).validate(anyMap());
            verify(mockOperation).get();
            verify(mockResultMapper).apply("success");
            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel." + OP_NAME), anyLong(), anyMap());
        }

        @Test
        @DisplayName("allowRetry=NULL → skips validation, executes operation, returns success")
        void allowRetryNull_skipsValidation_executesOperation() {
            when(mockOperation.get()).thenReturn("result");
            when(mockResultMapper.apply("result")).thenReturn(Map.of("data", "value"));

            ExtensionResponse response = execute(null);

            assertNotNull(response);
            verify(mockValidationOrchestrator, never()).validate(anyMap());
            verify(mockOperation).get();
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // PATH 2: allowRetry = TRUE  → validation passes → operation
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("When allowRetry=TRUE and validation passes")
    class AllowRetryEnabled_ValidationPasses {

        @Test
        @DisplayName("Should call validate, execute operation, record success, skip confirmation")
        void validationPasses_executesOperation() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            when(mockValidationOrchestrator.validate(vMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenReturn("operationResult");
            when(mockResultMapper.apply("operationResult")).thenReturn(Map.of("key", "value"));

            ExtensionResponse response = execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            assertNotNull(response);
            verify(mockValidationOrchestrator).validate(vMap);
            verify(mockOperation).get();
            verify(mockResultMapper).apply("operationResult");
            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel." + OP_NAME), anyLong(), anyMap());
            verify(mockStateManager, never()).put(anyString(), anyMap());
            verify(mockResponseGenerator, never()).generateConfirmationResponse(
                    any(), anyList(), anyString(), anyMap());
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // PATH 3: allowRetry = TRUE → validation fails → confirmation flow
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("When allowRetry=TRUE and validation fails")
    class AllowRetryEnabled_ValidationFails {

        @Test
        @DisplayName("Single validation error → stores state, returns confirmation response")
        void singleValidationError_returnsConfirmation() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            List<ValidationOrchestrator.ValidationResult> results = singleValidationError();
            ExtensionResponse expectedResponse = mock(ExtensionResponse.class);

            when(mockValidationOrchestrator.validate(vMap)).thenReturn(results);
            when(mockStateMapper.apply(anyString())).thenReturn(Map.of("stateId", "id"));
            when(mockResponseGenerator.generateConfirmationResponse(
                    any(ExtensionResponse.Error.ExceptionType.class),
                    anyList(), eq(CONFIRM_ACTION), anyMap()
            )).thenReturn(expectedResponse);

            ExtensionResponse response = execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            assertEquals(expectedResponse, response);
            verify(mockStateManager).put(anyString(), anyMap());
            verify(mockOperation, never()).get();
            verify(mockResponseGenerator, never()).generateValidationErrorResponse(anyList());
            verify(mockTelemetryHelper).recordValidationError(
                    eq("office365Excel." + OP_NAME), anyLong(), anyString(), anyMap());
        }

        @Test
        @DisplayName("Multiple validation errors → stores all results in state map")
        void multipleValidationErrors_storesAll() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            List<ValidationOrchestrator.ValidationResult> results = multipleValidationErrors();
            ExtensionResponse expectedResponse = mock(ExtensionResponse.class);

            when(mockValidationOrchestrator.validate(vMap)).thenReturn(results);
            when(mockStateMapper.apply(anyString())).thenReturn(Map.of("stateId", "id"));
            when(mockResponseGenerator.generateConfirmationResponse(
                    any(), anyList(), eq(CONFIRM_ACTION), anyMap()
            )).thenReturn(expectedResponse);

            ExtensionResponse response = execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            assertEquals(expectedResponse, response);
            verify(mockStateManager).put(anyString(), argThat(state ->
                    state.containsKey(SubCatalogConstants.VALIDATION_RESULTS)));
        }

        @Test
        @DisplayName("Confirmation response uses INPUT_ERROR exception type")
        void confirmationResponse_usesInputErrorType() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            List<ValidationOrchestrator.ValidationResult> results = singleValidationError();
            when(mockValidationOrchestrator.validate(vMap)).thenReturn(results);
            when(mockStateMapper.apply(anyString())).thenReturn(Map.of());
            when(mockResponseGenerator.generateConfirmationResponse(
                    any(), anyList(), anyString(), anyMap()
            )).thenReturn(mock(ExtensionResponse.class));

            execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            verify(mockResponseGenerator).generateConfirmationResponse(
                    eq(ExtensionResponse.Error.ExceptionType.INPUT_ERROR),
                    eq(results),
                    eq(CONFIRM_ACTION),
                    anyMap());
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // PATH 4: allowRetry = TRUE → validation throws IllegalStateException
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("When validation throws IllegalStateException (403/access denied)")
    class ValidationThrowsIllegalStateException {

        @Test
        @DisplayName("Returns access error response with remediation action")
        void illegalState_returnsAccessError() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            IllegalStateException exception = new IllegalStateException("Access denied to workbook");

            when(mockValidationOrchestrator.validate(vMap)).thenThrow(exception);

            ExtensionResponse response = execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            assertNotNull(response);
            verify(mockOperation, never()).get();
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OP_NAME), anyLong(), eq(exception), anyMap());
        }

        @Test
        @DisplayName("IllegalStateException with null message → still returns error response")
        void illegalState_nullMessage_returnsAccessError() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            IllegalStateException exception = new IllegalStateException();

            when(mockValidationOrchestrator.validate(vMap)).thenThrow(exception);

            ExtensionResponse response = execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            assertNotNull(response);
            verify(mockOperation, never()).get();
        }

        @Test
        @DisplayName("Should record LOGIC_ERROR + IllegalStateException in telemetry tags")
        void illegalState_recordsGeneralErrorTelemetry() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            IllegalStateException exception = new IllegalStateException("Infrastructure down");

            when(mockValidationOrchestrator.validate(vMap)).thenThrow(exception);

            execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OP_NAME), anyLong(), eq(exception),
                    argThat(tags ->
                            "LOGIC_ERROR".equals(tags.get("errorType")) &&
                            "IllegalStateException".equals(tags.get("errorClass"))
                    ));
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // PATH 5: Operation throws MustAuthorizeException
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("When operation throws MustAuthorizeException")
    class OperationThrowsMustAuthorizeException {

        @Test
        @DisplayName("invokeAsUser=false → returns authorization error response")
        void invokeAsUserFalse_returnsAuthErrorResponse() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            MustAuthorizeException authException = new MustAuthorizeException("Token expired");

            when(mockValidationOrchestrator.validate(vMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            ExtensionResponse response = execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            assertNotNull(response);
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OP_NAME), anyLong(), eq(authException), anyMap());
        }

        @Test
        @DisplayName("invokeAsUser=true → rethrows MustAuthorizeException")
        void invokeAsUserTrue_rethrowsException() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            MustAuthorizeException authException = new MustAuthorizeException("Auth required");

            when(mockValidationOrchestrator.validate(vMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(true);

            assertThrows(MustAuthorizeException.class, () ->
                    execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE));
        }

        @Test
        @DisplayName("invokeAsUser=false with null exception message → handles gracefully")
        void nullMessage_returnsAuthErrorResponse() {
            when(mockOperation.get()).thenThrow(new MustAuthorizeException((String) null));
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            ExtensionResponse response = execute(Boolean.FALSE);

            assertNotNull(response);
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OP_NAME), anyLong(),
                    any(MustAuthorizeException.class), anyMap());
        }

        @Test
        @DisplayName("Telemetry tags should have errorType=AUTHORIZATION, errorClass, errorMessage")
        void authError_setsCorrectTelemetryTags() {
            MustAuthorizeException authException = new MustAuthorizeException("token expired");

            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            execute(Boolean.FALSE);

            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OP_NAME), anyLong(), eq(authException),
                    argThat(tags ->
                            "AUTHORIZATION".equals(tags.get("errorType")) &&
                            "MustAuthorizeException".equals(tags.get("errorClass")) &&
                            "token expired".equals(tags.get("errorMessage"))
                    ));
        }

        @Test
        @DisplayName("Null exception message → telemetry defaults to 'Authorization required'")
        void authErrorNullMessage_defaultsTelemetryMessage() {
            MustAuthorizeException authException = new MustAuthorizeException((String) null);

            when(mockOperation.get()).thenThrow(authException);
            when(mockRequestContext.invokeAsUser()).thenReturn(false);

            execute(Boolean.FALSE);

            verify(mockTelemetryHelper).recordError(
                    anyString(), anyLong(), any(),
                    argThat(tags -> "Authorization required".equals(tags.get("errorMessage"))));
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // PATH 6: Operation throws generic Exception
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("When operation throws generic Exception")
    class OperationThrowsGenericException {

        @Test
        @DisplayName("Returns error response and removes state")
        void genericException_returnsErrorAndRemovesState() {
            RuntimeException exception = new RuntimeException("Graph API error");

            when(mockOperation.get()).thenThrow(exception);

            ExtensionResponse response = execute(Boolean.FALSE);

            assertNotNull(response);
            verify(mockStateManager).remove(anyString());
            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OP_NAME), anyLong(), eq(exception), anyMap());
        }

        @Test
        @DisplayName("Exception with null message → handles gracefully")
        void genericExceptionNullMessage_returnsError() {
            RuntimeException exception = new RuntimeException((String) null);

            when(mockOperation.get()).thenThrow(exception);

            ExtensionResponse response = execute(Boolean.FALSE);

            assertNotNull(response);
            verify(mockStateManager).remove(anyString());
        }

        @Test
        @DisplayName("allowRetry=TRUE, validation passes, operation fails → removes state")
        void afterValidationPasses_operationFails_removesState() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            RuntimeException exception = new RuntimeException("Unexpected error");

            when(mockValidationOrchestrator.validate(vMap)).thenReturn(Collections.emptyList());
            when(mockOperation.get()).thenThrow(exception);

            ExtensionResponse response = execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            assertNotNull(response);
            verify(mockValidationOrchestrator).validate(vMap);
            verify(mockStateManager).remove(anyString());
        }

        @Test
        @DisplayName("Telemetry tags should have errorType=LOGIC_ERROR, errorClass, errorMessage")
        void generalError_setsCorrectTelemetryTags() {
            RuntimeException exception = new RuntimeException("Something went wrong");

            when(mockOperation.get()).thenThrow(exception);

            execute(Boolean.FALSE);

            verify(mockTelemetryHelper).recordError(
                    eq("office365Excel." + OP_NAME), anyLong(), eq(exception),
                    argThat(tags ->
                            "LOGIC_ERROR".equals(tags.get("errorType")) &&
                            "RuntimeException".equals(tags.get("errorClass")) &&
                            "Something went wrong".equals(tags.get("errorMessage"))
                    ));
        }

        @Test
        @DisplayName("Null exception message → telemetry defaults to 'Unknown error'")
        void generalErrorNullMessage_defaultsTelemetryMessage() {
            RuntimeException exception = new RuntimeException((String) null);

            when(mockOperation.get()).thenThrow(exception);

            execute(Boolean.FALSE);

            verify(mockTelemetryHelper).recordError(
                    anyString(), anyLong(), any(),
                    argThat(tags -> "Unknown error".equals(tags.get("errorMessage"))));
        }
    }

    // ═══════════════════════════════════════════════════════════════════
    // TELEMETRY HELPER METHODS & EDGE CASES
    // ═══════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Telemetry helpers and edge cases")
    class TelemetryAndEdgeCases {

        @Test
        @DisplayName("prepareTelemetryTags: null telemetryTags → defaults to empty map, no NPE")
        void nullTelemetryTags_shouldNotThrow() {
            when(mockOperation.get()).thenReturn("result");
            when(mockResultMapper.apply("result")).thenReturn(Map.of("key", "val"));

            ExtensionResponse response = execute(OP_NAME, validationMap(), stateValues(), null, Boolean.FALSE);

            assertNotNull(response);
            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel." + OP_NAME), anyLong(), anyMap());
        }

        @Test
        @DisplayName("prepareTelemetryTags: includes operation, validationResourceCount, stateValueCount")
        void prepareTags_includesCounts() {
            when(mockOperation.get()).thenReturn("ok");
            when(mockResultMapper.apply("ok")).thenReturn(Map.of());

            execute(OP_NAME, validationMap(), stateValues(), Map.of(), Boolean.FALSE);

            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel." + OP_NAME), anyLong(),
                    argThat(tags ->
                            "2".equals(tags.get("validationResourceCount")) &&
                            "2".equals(tags.get("stateValueCount")) &&
                            OP_NAME.equals(tags.get("operation"))
                    ));
        }

        @Test
        @DisplayName("prepareTelemetryTags: preserves custom tags from caller")
        void prepareTags_preservesCustomTags() {
            when(mockOperation.get()).thenReturn("ok");
            when(mockResultMapper.apply("ok")).thenReturn(Map.of());

            Map<String, String> custom = new HashMap<>();
            custom.put("myTag", "myValue");

            execute(OP_NAME, validationMap(), stateValues(), custom, Boolean.FALSE);

            verify(mockTelemetryHelper).recordSuccess(
                    anyString(), anyLong(),
                    argThat(tags -> "myValue".equals(tags.get("myTag"))));
        }

        @Test
        @DisplayName("baseMetric is built as 'office365Excel.' + operationName")
        void baseMetric_includesOperationName() {
            String customOp = "getCellData";
            when(mockOperation.get()).thenReturn("data");
            when(mockResultMapper.apply("data")).thenReturn(Map.of("cell", "A1"));

            execute(customOp, validationMap(), stateValues(), telemetryTags(), Boolean.FALSE);

            verify(mockTelemetryHelper).recordSuccess(
                    eq("office365Excel.getCellData"), anyLong(), anyMap());
        }

        @Test
        @DisplayName("recordValidationFailure: records error count, failed fields, and error message")
        void validationFailure_recordsCountAndFields() {
            Map<Validator.ValidationResource, String> vMap = validationMap();
            List<ValidationOrchestrator.ValidationResult> results = multipleValidationErrors();
            ExtensionResponse expectedResponse = mock(ExtensionResponse.class);

            when(mockValidationOrchestrator.validate(vMap)).thenReturn(results);
            when(mockStateMapper.apply(anyString())).thenReturn(Map.of());
            when(mockResponseGenerator.generateConfirmationResponse(
                    any(), anyList(), anyString(), anyMap()
            )).thenReturn(expectedResponse);

            execute(OP_NAME, vMap, stateValues(), telemetryTags(), Boolean.TRUE);

            verify(mockTelemetryHelper).recordValidationError(
                    eq("office365Excel." + OP_NAME), anyLong(),
                    argThat(msg -> msg.contains("2 errors")),
                    argThat(tags ->
                            "2".equals(tags.get("validationErrorCount")) &&
                            tags.get("validationErrors") != null
                    ));
        }
    }
}
