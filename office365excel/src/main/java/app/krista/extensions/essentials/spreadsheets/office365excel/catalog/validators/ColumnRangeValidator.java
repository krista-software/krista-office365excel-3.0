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
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;

import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.COLUMN_RANGE;

/**
 * Validator for Column Range in Office365Excel extension.
 */
public class ColumnRangeValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(ColumnRangeValidator.class);

    @Override
    public Boolean validate(String columnRange, Map<ValidationResource, String> context) {
        if (columnRange == null || columnRange.trim().isEmpty()) {
            LOGGER.debug("Column range validation failed: column range is null or empty");
            return false;
        }

        try {
            if (!Worksheets.isValidColumnRange(columnRange)) {
                LOGGER.debug("Column range validation failed: '{}' is not a valid Excel column range format", columnRange);
                return false;
            }
            
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating column range: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return COLUMN_RANGE;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid Excel column range (e.g., A:C).";
    }

    @Override
    public String getConfirmStepMessage(String columnRange, Map<ValidationResource, String> context) {
        return "The column range '" + columnRange + "' is not valid. Would you like to enter a different column range?";
    }

    @Override
    public String getErrMessage(String columnRange) {
        return "Invalid column range: " + columnRange + ". Please provide a valid Excel column range in the format A:C.";
    }
}