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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;
import java.util.regex.Pattern;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.FORMULA_NAME;

/**
 * Validator for Excel formula names in Office365Excel extension.
 */
public class FormulaNameValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(FormulaNameValidator.class);
    private static final int MAX_NAME_LENGTH = 255; // Excel's limit for named ranges
    private static final Pattern VALID_NAME_PATTERN = Pattern.compile("^[a-zA-Z_][a-zA-Z0-9_.]*$");

    @Override
    public Boolean validate(String formulaName, Map<ValidationResource, String> context) {
        if (formulaName == null || formulaName.trim().isEmpty()) {
            LOGGER.debug("Formula name validation failed: name is null or empty");
            return false;
        }
        
        try {
            // Check length constraint
            if (formulaName.length() > MAX_NAME_LENGTH) {
                LOGGER.debug("Formula name validation failed: name exceeds maximum length");
                return false;
            }
            
            // Check for valid characters (must start with letter or underscore, can contain letters, numbers, underscores, periods)
            if (!VALID_NAME_PATTERN.matcher(formulaName).matches()) {
                LOGGER.debug("Formula name validation failed: name contains invalid characters");
                return false;
            }
            
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating formula name: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return FORMULA_NAME;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid formula name. Names must start with a letter or underscore and can contain letters, numbers, underscores, and periods.";
    }

    @Override
    public String getConfirmStepMessage(String formulaName, Map<ValidationResource, String> context) {
        return "The formula name '" + formulaName + "' is invalid. Would you like to enter a different name?";
    }

    @Override
    public String getErrMessage(String formulaName) {
        return "Invalid formula name: '" + formulaName + "'. Names must start with a letter or underscore and can contain letters, numbers, underscores, and periods.";
    }
}