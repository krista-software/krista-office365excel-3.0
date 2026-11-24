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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.COLUMN_INDEX;

/**
 * Validator for Column Index in Office365Excel extension.
 * Ensures that column index is a positive number greater than zero.
 */
public class ColumnIndexValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(ColumnIndexValidator.class);

    @Override
    public Boolean validate(String columnIndexStr, Map<ValidationResource, String> context) {
        if (columnIndexStr == null || columnIndexStr.trim().isEmpty()) {
            LOGGER.debug("Column index validation failed: value is null or empty");
            return false;
        }

        try {
            double columnIndex = Double.parseDouble(columnIndexStr);
            if (columnIndex <= 0) {
                LOGGER.debug("Column index validation failed: value must be greater than 0, got {}", columnIndex);
                return false;
            }

            return true;
        } catch (NumberFormatException cause) {
            LOGGER.debug("Column index validation failed: value is not a valid number: {}", columnIndexStr);
            return false;
        } catch (Exception cause) {
            LOGGER.error("Error validating column index: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return COLUMN_INDEX;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.NUMBER_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid column index (must be greater than 0).";
    }

    @Override
    public String getConfirmStepMessage(String columnIndex, Map<ValidationResource, String> context) {
        return "The column index '" + columnIndex + "' is invalid. Column index must be greater than 0. Would you like to enter a different value?";
    }

    @Override
    public String getErrMessage(String columnIndex) {
        return "Invalid column index: " + columnIndex + ". Please provide a number greater than 0.";
    }
}