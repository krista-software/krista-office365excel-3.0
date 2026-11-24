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
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Worksheets;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.RANGE;

public class RangeValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(RangeValidator.class);

    @Override
    public Boolean validate(String range, Map<ValidationResource, String> context) {
        if (range == null || range.trim().isEmpty()) {
            LOGGER.debug("Range validation failed: range is null or empty");
            return false;
        }

        try {
            if (!Worksheets.isValidRange(range)) {
                LOGGER.debug("Range validation failed: '{}' is not a valid Excel range format", range);
                return false;
            }
            
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating range: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return RANGE;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid Excel range using the format A1:B10 (single cell: A1, multiple cells: A1:C5, entire column: A:A, entire row: 1:1).";
    }

    @Override
    public String getConfirmStepMessage(String range, Map<ValidationResource, String> context) {
        return "The range '" + range + "' is not in a valid Excel format. Please use formats like A1:B10 for cell ranges, A1 for single cells, A:A for columns, or 1:1 for rows. Would you like to try again?";
    }

    @Override
    public String getErrMessage(String range) {
        return "Invalid Excel range '" + range + "'. Please use valid Excel range formats like A1:B10 (cell range), A1 (single cell), A:A (column), or 1:1 (row).";
    }
}