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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.CELL_ID;

/**
 * Validator for Cell ID in Office365Excel extension.
 * Ensures that cell ID is in valid Excel format (e.g., A1, B2, etc.).
 */
public class CellIdValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(CellIdValidator.class);

    @Override
    public Boolean validate(String cellId, Map<ValidationResource, String> context) {
        if (cellId == null || cellId.trim().isEmpty()) {
            LOGGER.debug("Cell ID validation failed: value is null or empty");
            return false;
        }
        
        try {
            if (!Worksheets.isValidCellId(cellId)) {
                LOGGER.debug("Cell ID validation failed: '{}' is not a valid Excel cell ID", cellId);
                return false;
            }
            
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating cell ID: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return CELL_ID;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid Excel cell reference using column letter(s) followed by row number (e.g., A1, B2, AA10, XFD1048576).";
    }

    @Override
    public String getConfirmStepMessage(String cellId, Map<ValidationResource, String> context) {
        return "The cell reference '" + cellId + "' is not valid. Excel cell references must have column letter(s) followed by a row number (e.g., A1, B2, AA10). Would you like to try again?";
    }

    @Override
    public String getErrMessage(String cellId) {
        return "Invalid Excel cell reference '" + cellId + "'. Please use the format: column letter(s) + row number (e.g., A1, B2, AA10).";
    }
}