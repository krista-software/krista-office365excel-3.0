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
import app.krista.extension.executor.ExtensionResponseBuilder;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.ExtensionResponseFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.NamedFieldFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.RemediationActionFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.model.field.NamedField;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

@Service
public class ExtensionResponseGenerator {

    public static final String ERROR_MESSAGE = "ErrorMessage";
    public static final String STEP_MESSAGE = "StepMessage";
    public static final String FIELD = "Field";
    private static final Logger LOGGER = LoggerFactory.getLogger(ExtensionResponseGenerator.class);

    /**
     * Generates a confirmation response for the first step of the 3-step validation flow.
     * Asks the user to confirm whether they want to re-enter corrected values.
     *
     * @param exceptionType the type of exception that triggered the validation
     * @param validationResults list of validation results containing error details
     * @param subCatalogRequestName the name of the catalog request being executed
     * @param state the current state to be preserved across validation steps
     * @return ExtensionResponse with an Ask remediation action containing a switch field for re-entry confirmation
     */
    public ExtensionResponse generateConfirmationResponse(
            ExtensionResponse.Error.ExceptionType exceptionType,
            List<ValidationOrchestrator.ValidationResult> validationResults,
            String subCatalogRequestName,
            Map<String, Object> state) {

        List<NamedField> fields = List.of(NamedFieldFactory.createSwitchField(SubCatalogConstants.REENTER));
        Map<String, Object> stringStringMap = generateResponse(validationResults, false);
        return ExtensionResponseFactory.create((String) stringStringMap.get(ERROR_MESSAGE), exceptionType,
                List.of(RemediationActionFactory.createAskAction((String) stringStringMap.get(STEP_MESSAGE), fields)),
                subCatalogRequestName, state);
    }

    /**
     * Internal helper method to generate response messages and fields from validation results.
     *
     * @param validationResults list of validation results to process
     * @param fetchResponse true to generate fetch step messages and fields, false for confirmation step
     * @return map containing ERROR_MESSAGE, STEP_MESSAGE, and FIELD entries
     */
    private Map<String, Object> generateResponse(List<ValidationOrchestrator.ValidationResult> validationResults, boolean fetchResponse) {
        StringBuilder stepMessage = new StringBuilder();
        StringBuilder errMessage = new StringBuilder();
        List<NamedField> fields = new ArrayList<>();
        for (ValidationOrchestrator.ValidationResult validationResult : validationResults) {
            stepMessage.append(!fetchResponse ? validationResult.getConfirmStepMessage() : validationResult.getFetchStepMessage()).append("\n");
            errMessage.append(validationResult.getErrMessage()).append("\n");
            if (fetchResponse) {

                //If Field Type is File
                fields.add(NamedFieldFactory.createFieldWithAttributes(validationResult.getFetchFieldName(), validationResult.getFieldType(), null));
            }
        }
        return Map.of(STEP_MESSAGE, stepMessage.toString(), ERROR_MESSAGE, errMessage.toString(), FIELD, fields);
    }

    /**
     * Generates a fetch response for the second step of the 3-step validation flow.
     * Asks the user to provide corrected values for the invalid fields.
     *
     * @param exceptionType the type of exception that triggered the validation
     * @param validationResults list of validation results containing error details and field information
     * @param subCatalogRequestName the name of the catalog request being executed
     * @param state the current state to be preserved across validation steps
     * @return ExtensionResponse with an Ask remediation action containing input fields for corrected values
     */
    @SuppressWarnings("unchecked")
    public ExtensionResponse generateFetchResponse(ExtensionResponse.Error.ExceptionType exceptionType,
                                                   List<ValidationOrchestrator.ValidationResult> validationResults,
                                                   String subCatalogRequestName,
                                                   Map<String, Object> state) {

        LOGGER.debug("Validation failed for : {}", validationResults.size());
        Map<String, Object> stringStringMap = generateResponse(validationResults, true);
        ExtensionResponse extensionResponse = ExtensionResponseFactory.create((String) stringStringMap.get(ERROR_MESSAGE), exceptionType,
                List.of(
                        RemediationActionFactory.createAskAction((String) stringStringMap.get(STEP_MESSAGE), (List<NamedField>) stringStringMap.get(FIELD))
                ),
                subCatalogRequestName, state);
        LOGGER.debug("Extension Response : {}", extensionResponse.getState());
        return extensionResponse;
    }

    /**
     * Generates a deny response when the user does not provide corrected values in the fetch step.
     * Informs all participants that the required values were not provided.
     *
     * @param exceptionType the type of exception that triggered the validation
     * @param validationResults list of validation results containing field names that were not provided
     * @param subCatalogRequestName the name of the catalog request being executed
     * @param state the current state to be preserved
     * @return ExtensionResponse with an Inform remediation action notifying about missing values
     */
    public ExtensionResponse generateFetchDenyResponse( ExtensionResponse.Error.ExceptionType exceptionType,
                                                        List<ValidationOrchestrator.ValidationResult> validationResults,
                                                        String subCatalogRequestName,
                                                        Map<String, Object> state) {
        StringJoiner stepMessage = new StringJoiner(", ", "Updated value for ", " were not provided.");
        StringBuilder errMessage = new StringBuilder();
        for (ValidationOrchestrator.ValidationResult validationResult : validationResults) {
            stepMessage.add(String.format("'%s'", validationResult.getFetchFieldName()));
            errMessage.append(validationResult.getErrMessage()).append(System.lineSeparator());
        }
        return ExtensionResponseFactory.create(errMessage.toString(), exceptionType,
                List.of(RemediationActionFactory.createInformActionALLParticipants(stepMessage.toString(), List.of())),
                subCatalogRequestName, state);
    }

    /**
     * Generates a direct validation error response (WITHOUT retry).
     * Returns a simple error message without remediation actions.
     * Used when the "Allow Retry" checkbox is unchecked or null.
     *
     * @param validationResults list of validation results containing error messages
     * @return ExtensionResponse with a simple failure message without remediation actions
     */
    public ExtensionResponse generateValidationErrorResponse(
            List<ValidationOrchestrator.ValidationResult> validationResults) {

        LOGGER.debug("Generating direct validation error response for {} validation failures", validationResults.size());

        // Build error message from validation results
        StringBuilder errMessage = new StringBuilder();
        for (ValidationOrchestrator.ValidationResult result : validationResults) {
            errMessage.append(result.getErrMessage()).append("\n");
        }

        // Return simple failure response without remediation
        return new ExtensionResponseBuilder()
                .failure(errMessage.toString().trim())
                .build();
    }

}
