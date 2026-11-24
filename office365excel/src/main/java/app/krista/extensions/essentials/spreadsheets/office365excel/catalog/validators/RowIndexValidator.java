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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.ROW_INDEX;

/**
 * Validator for Row Index in Office365Excel extension.
 * Ensures that row index is a valid positive number.
 */
public class RowIndexValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowIndexValidator.class);

    @Override
    public Boolean validate(String rowIndexStr, Map<ValidationResource, String> context) {
        if (rowIndexStr == null || rowIndexStr.trim().isEmpty()) {
            LOGGER.debug("Row index validation failed: value is null or empty");
            return false;
        }
        
        try {
            double rowIndex = Double.parseDouble(rowIndexStr);
            if (rowIndex <= 0) {
                LOGGER.debug("Row index validation failed: value must be greater than 0, got {}", rowIndex);
                return false;
            }
            
            return true;
        } catch (NumberFormatException cause) {
            LOGGER.debug("Row index validation failed: value is not a valid number: {}", rowIndexStr);
            return false;
        } catch (Exception cause) {
            LOGGER.error("Error validating row index: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return ROW_INDEX;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.NUMBER_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid row number (must be a positive number starting from 1). Row 1 is the first row in Excel.";
    }

    @Override
    public String getConfirmStepMessage(String rowIndex, Map<ValidationResource, String> context) {
        return "The row number '" + rowIndex + "' is not valid. Please enter a positive number (1 or greater) representing the Excel row. Would you like to try again?";
    }

    @Override
    public String getErrMessage(String rowIndex) {
        return "Invalid row number '" + rowIndex + "'. Please provide a positive number (1 or greater) representing the Excel row.";
    }
}