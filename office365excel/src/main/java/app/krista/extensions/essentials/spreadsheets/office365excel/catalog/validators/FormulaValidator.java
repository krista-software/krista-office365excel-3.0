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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.FORMULA;

/**
 * Validator for Excel formulas in Office365Excel extension.
 */
public class FormulaValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(FormulaValidator.class);
    private static final int MAX_FORMULA_LENGTH = 8192; // Excel's limit for formula length

    @Override
    public Boolean validate(String formula, Map<ValidationResource, String> context) {
        if (formula == null || formula.trim().isEmpty()) {
            LOGGER.debug("Formula validation failed: formula is null or empty");
            return false;
        }
        
        try {
            // Check length constraint
            if (formula.length() > MAX_FORMULA_LENGTH) {
                LOGGER.debug("Formula validation failed: formula exceeds maximum length");
                return false;
            }
            
            // Basic formula validation - must start with =
            if (!formula.startsWith("=")) {
                LOGGER.debug("Formula validation failed: formula does not start with =");
                return false;
            }
            
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating formula: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return FORMULA;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid Excel formula starting with '=' (e.g., =SUM(A1:A10), =IF(B1>0,\"Positive\",\"Negative\"), =VLOOKUP(C1,A:B,2,FALSE)).";
    }

    @Override
    public String getConfirmStepMessage(String formula, Map<ValidationResource, String> context) {
        return "The formula '" + formula + "' is not valid. Excel formulas must start with '=' and cannot exceed 8,192 characters. Please check your formula syntax. Would you like to try again?";
    }

    @Override
    public String getErrMessage(String formula) {
        return "Invalid Excel formula '" + formula + "'. Formulas must start with '=' and be valid Excel expressions (max 8,192 characters).";
    }
}