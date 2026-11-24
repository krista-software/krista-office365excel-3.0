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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.COLUMN_NAME;

/**
 * Validator for Column Names in Office365Excel extension.
 * Ensures that column names exist in the specified table.
 */
public class ColumnNameValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(ColumnNameValidator.class);
    private final WorkbookService workbookService;

    public ColumnNameValidator(WorkbookService workbookService) {
        this.workbookService = workbookService;
    }

    @Override
    public Boolean validate(String columnName, Map<ValidationResource, String> context) {
        if (columnName == null || columnName.trim().isEmpty()) {
            LOGGER.debug("Column name validation failed: name is null or empty");
            return false;
        }

        try {
            String driveId = context.get(ValidationResource.DRIVE_ID);
            String fileId = context.get(ValidationResource.FILE_ID);
            String tableName = context.get(ValidationResource.TABLE_NAME);
            
            if (driveId == null || fileId == null || tableName == null) {
                LOGGER.debug("Column name validation failed: missing required context (driveId, fileId, or tableName)");
                return false;
            }
            
            // Check if the column exists in the table
            return workbookService.isColumnExistsInTable(driveId, fileId, tableName, columnName);
        } catch (Exception cause) {
            LOGGER.error("Error validating column name: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return COLUMN_NAME;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter the exact name of a column that exists in the table. Column names are case-sensitive and must match exactly as they appear in Excel.";
    }

    @Override
    public String getConfirmStepMessage(String columnName, Map<ValidationResource, String> context) {
        return "The column '" + columnName + "' doesn't exist in this table. Please check the spelling and capitalization, as column names are case-sensitive. Would you like to try a different column name?";
    }

    @Override
    public String getErrMessage(String columnName) {
        return "Column '" + columnName + "' not found in the table. Please verify the column name exists and is spelled correctly (case-sensitive).";
    }
}