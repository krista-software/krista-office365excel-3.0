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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.ROWS;

/**
 * Validator for Row Data in Office365Excel extension.
 */
public class RowDataValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowDataValidator.class);
    private static final int MAX_ROW_LENGTH = 16384; // Excel's maximum column count

    @Override
    public Boolean validate(String rowData, Map<ValidationResource, String> context) {
        if (rowData == null || rowData.isBlank()) {
            return false;
        }
        
        try {
            return rowData.length() <= MAX_ROW_LENGTH;
        } catch (Exception cause) {
            LOGGER.error("Error validating row data: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return ROWS;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.LIST_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter valid row data.";
    }

    @Override
    public String getConfirmStepMessage(String rowData, Map<ValidationResource, String> context) {
        return "The row data is invalid. Would you like to enter different row data?";
    }

    @Override
    public String getErrMessage(String rowData) {
        return "The row data is invalid. Please ensure it's in the correct format and doesn't exceed the maximum length.";
    }
}