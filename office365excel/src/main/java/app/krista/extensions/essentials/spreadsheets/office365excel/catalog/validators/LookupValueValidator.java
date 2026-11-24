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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.LOOKUP_VALUE;

/**
 * Validator for lookup values in Office365Excel extension.
 * Ensures that lookup value is not empty.
 */
public class LookupValueValidator implements Validator {
    private static final Logger LOGGER = LoggerFactory.getLogger(LookupValueValidator.class);
    private static final int MAX_LOOKUP_VALUE_LENGTH = 255; // Reasonable limit for lookup values

    @Override
    public Boolean validate(String lookupValue, Map<ValidationResource, String> context) {
        if (lookupValue == null || lookupValue.trim().isEmpty()) {
            LOGGER.debug("Lookup value validation failed: value is null or empty");
            return false;
        }

        try {
            if (lookupValue.length() > MAX_LOOKUP_VALUE_LENGTH) {
                LOGGER.debug("Lookup value validation failed: value exceeds maximum length of {} characters", MAX_LOOKUP_VALUE_LENGTH);
                return false;
            }

            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating lookup value: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return LOOKUP_VALUE;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid lookup value.";
    }

    @Override
    public String getConfirmStepMessage(String lookupValue, Map<ValidationResource, String> context) {
        if (lookupValue == null || lookupValue.trim().isEmpty()) {
            return "The lookup value cannot be empty. Would you like to enter a lookup value?";
        } else {
            return "The lookup value is too long. Would you like to enter a shorter lookup value?";
        }
    }

    @Override
    public String getErrMessage(String lookupValue) {
        if (lookupValue == null || lookupValue.trim().isEmpty()) {
            return "Lookup value cannot be empty.";
        } else {
            return "Lookup value is too long. Please limit your lookup value to " + MAX_LOOKUP_VALUE_LENGTH + " characters.";
        }
    }
}