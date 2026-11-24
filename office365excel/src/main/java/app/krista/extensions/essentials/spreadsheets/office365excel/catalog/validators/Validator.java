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

import java.util.Map;

/**
 * Interface for all validators in the Office365Excel extension.
 * Each validator is responsible for validating a specific type of resource.
 */
public interface Validator {

    enum ValidationResource {
        DRIVE_ID,
        FILE_ID,
        SHEET_NAME,
        TABLE_NAME,
        ROW_DATA,
        CSV_FILE,
        RANGE,
        COLUMN_INDEX,
        CELL_ID,
        ROW_INDEX,
        COLUMN_NAME,
        FORMULA,
        FORMULA_NAME,
        COMMENT,
        SHARED_FILE_NAME,
        KEY_ROWS,
        LOOKUP_COLUMNS,
        COLUMNS,
        LOOKUP_ROWS,
        LOOKUP_VALUE,
        PATH,
        NAME,
        EXCEL_FILE,
        COLUMN_RANGE,
        KEY_COLUMN_NAME,
        WORKBOOK_ID

    }

    /**
     * Validates the given resource against validation rules
     * @param resourceValue The resource value to validate
     * @param context Additional context for validation
     * @return true if valid, false otherwise
     */
    Boolean validate(String resourceValue, Map<ValidationResource, String> context);

    /**
     * Gets the field name for fetch operations
     */
    String getFetchFieldName();

    /**
     * Gets the field type for the UI
     */
    String getFieldType();

    /**
     * Gets the message to display during fetch step
     */
    String getFetchStepMessage();

    /**
     * Gets the confirmation message after validation
     */
    String getConfirmStepMessage(String resourceValue, Map<ValidationResource, String> context);

    /**
     * Gets the error message for invalid resources
     */
    String getErrMessage(String resourceValue);
}
