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
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.AuthorizationExceptionHandler;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.ExtensionResponseFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.RemediationActionFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.Validator;
import app.krista.ksdk.context.RequestContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.function.Function;
import java.util.function.Supplier;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.*;

/**
 * Generic handler for Excel operations with validation and error handling
 *
 * @param <T> The return type of the operation
 */
public class AreaOperationHandler<T> {
    private static final Logger LOGGER = LoggerFactory.getLogger(AreaOperationHandler.class);

    private final ValidationOrchestrator validationOrchestrator;
    private final ErrorHandlingStateManager internalStateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final RequestContext requestContext;
    private final TelemetryHelper telemetryHelper;

    public AreaOperationHandler(ValidationOrchestrator validationOrchestrator, ErrorHandlingStateManager internalStateManager,
                                     ExtensionResponseGenerator responseGenerator, RequestContext requestContext,
                                     TelemetryHelper telemetryHelper) {
        this.validationOrchestrator = validationOrchestrator;
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
        this.requestContext = requestContext;
        this.telemetryHelper = telemetryHelper;
    }

    /**
     * Executes an Excel operation with validation and error handling
     *
     * @param operationName       Name of the operation for logging
     * @param validationMap       Map of validation resources to validate
     * @param stateValues         Map of values to store in state for reentry
     * @param confirmationAction  Subcatalog constant for confirmation action
     * @param stateMapperFunction Function to create state map for confirmation
     * @param operation           The operation to execute if validation passes
     * @param resultMapper        Function to map the operation result to response map
     * @param allowRetry          Whether to allow retry on validation failure (null or false = direct error)
     * @return Extension response with result or error information
     */
    public ExtensionResponse execute(String operationName, Map<Validator.ValidationResource, String> validationMap,
                                     Map<String, Object> stateValues, String confirmationAction,
                                     Function<String, Map<String, Object>> stateMapperFunction,
                                     Supplier<T> operation, Function<T, Map<String, Object>> resultMapper,
                                     Map<String, String> telemetryTags, Boolean allowRetry) {

        long startTime = System.currentTimeMillis();
        String baseMetric = "office365Excel." + operationName;
        Map<String, String> allTags = prepareTelemetryTags(operationName, validationMap, stateValues, telemetryTags);

        LOGGER.info("{} start with parameters: {}", operationName, stateValues);
        String stateId = UUID.randomUUID().toString();
        List<ValidationOrchestrator.ValidationResult> validationResults;
        if (!Boolean.TRUE.equals(allowRetry)) {
            // Validation (including the worksheet existence API call) is only run when
            // allowRetry=true so the user can get an interactive retry prompt.
            // Without retry, skip validation and let the operation fail naturally with
            // the real Graph API error if inputs are wrong.
            LOGGER.debug("Skipping validation for {} because allowRetry is not enabled", operationName);
            validationResults = List.of();
        } else {
            try {
                validationResults = validationOrchestrator.validate(validationMap);
            } catch (IllegalStateException cause) {
                LOGGER.error("Access or infrastructure error during validation in {}: {}", operationName, cause.getMessage(), cause);
                recordGeneralError(baseMetric, startTime, cause, allTags);
                return ExtensionResponseFactory.create(
                        cause, cause.getMessage(),
                        ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                        List.of(RemediationActionFactory.createInformAction(
                                "An access error occurred. Please verify you have permission to access the workbook and try again.", List.of())),
                        null, null);
            }
        }

        if (!validationResults.isEmpty()) {
            LOGGER.debug("Validation failed for {} with {} errors", operationName, validationResults.size());

            recordValidationFailure(baseMetric, startTime, validationResults, allTags);

            // Check if retry is allowed
            if (Boolean.TRUE.equals(allowRetry)) {
                // Allow Retry is ENABLED - Generate interactive 3-step retry flow
                LOGGER.debug("Allow Retry is enabled - generating confirmation response for interactive retry");

                Map<String, Object> state = new HashMap<>(stateValues);
                state.put(SubCatalogConstants.VALIDATION_RESULTS, validationResults);
                internalStateManager.put(stateId, state);

                LOGGER.info("Stored validation state with ID: {}", stateId);

                return responseGenerator.generateConfirmationResponse(ExtensionResponse.Error.ExceptionType.INPUT_ERROR,
                        validationResults, confirmationAction, stateMapperFunction.apply(stateId));
            } else {
                // Allow Retry is DISABLED or null - Return direct error without retry
                LOGGER.debug("Allow Retry is disabled or null - generating direct validation error response");

                return responseGenerator.generateValidationErrorResponse(validationResults);
            }
        }

        try {
            T result = operation.get();
            LOGGER.info("Successfully executed {}", operationName);

            recordSuccess(baseMetric, startTime, allTags);

            return ExtensionResponseFactory.create(resultMapper.apply(result));
        } catch (MustAuthorizeException cause) {
            LOGGER.error("Authorization error in {}: {}", operationName, cause.getMessage(), cause);

            recordAuthorizationError(baseMetric, startTime, cause, allTags);

            return AuthorizationExceptionHandler.handleAuthorizationException(cause, requestContext.invokeAsUser());
        } catch (Exception cause) {
            LOGGER.error("Error in {}: {}", operationName, cause.getMessage(), cause);

            recordGeneralError(baseMetric, startTime, cause, allTags);

            internalStateManager.remove(stateId);
            return ExtensionResponseFactory.create(
                    cause, "Failed to execute " + operationName + ": " + cause.getMessage(),
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(RemediationActionFactory.createInformAction(
                            "An error occurred. Please try again.", List.of())),
                    null, null);
        }
    }

    // ===== TELEMETRY HELPER METHODS =====

    private Map<String, String> prepareTelemetryTags(String operationName,
                                                      Map<Validator.ValidationResource, String> validationMap,
                                                      Map<String, Object> stateValues,
                                                      Map<String, String> telemetryTags) {
        Map<String, String> allTags = new HashMap<>(telemetryTags != null ? telemetryTags : Map.of());
        allTags.put("operation", operationName);
        allTags.put("validationResourceCount", String.valueOf(validationMap.size()));
        allTags.put("stateValueCount", String.valueOf(stateValues.size()));
        return allTags;
    }

    private void recordValidationFailure(String baseMetric, long startTime,
                                       List<ValidationOrchestrator.ValidationResult> validationResults,
                                       Map<String, String> allTags) {
        allTags.put("validationErrorCount", String.valueOf(validationResults.size()));

        String failedFields = validationResults.stream()
            .map(ValidationOrchestrator.ValidationResult::getFetchFieldName)
            .distinct()
            .reduce((a, b) -> a + "," + b)
            .orElse("");

        allTags.put("validationErrors", failedFields);

        String validationErrorMessage = "Validation failed with " + validationResults.size() + " errors: " +
            validationResults.stream()
                .map(vr -> vr.getFetchFieldName() + "=" + vr.getErrMessage())
                .reduce((a, b) -> a + "; " + b)
                .orElse("");

        telemetryHelper.recordValidationError(baseMetric, startTime, validationErrorMessage, allTags);
    }

    private void recordSuccess(String baseMetric, long startTime, Map<String, String> allTags) {
        telemetryHelper.recordSuccess(baseMetric, startTime, allTags);
    }

    private void recordAuthorizationError(String baseMetric, long startTime, MustAuthorizeException cause,
                                        Map<String, String> allTags) {
        allTags.put("errorType", AUTHORIZATION);
        allTags.put("errorMessage", cause.getMessage() != null ? cause.getMessage() : "Authorization required");
        allTags.put("errorClass", cause.getClass().getSimpleName());
        telemetryHelper.recordError(baseMetric, startTime, cause, allTags);
    }

    private void recordGeneralError(String baseMetric, long startTime, Exception cause,
                                  Map<String, String> allTags) {
        allTags.put("errorType", LOGIC_ERROR);
        allTags.put("errorMessage", cause.getMessage() != null ? cause.getMessage() : "Unknown error");
        allTags.put("errorClass", cause.getClass().getSimpleName());
        telemetryHelper.recordError(baseMetric, startTime, cause, allTags);
    }
}