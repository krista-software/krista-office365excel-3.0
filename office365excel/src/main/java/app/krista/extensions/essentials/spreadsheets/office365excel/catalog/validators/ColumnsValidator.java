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

import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.COLUMNS;


public class ColumnsValidator implements Validator {
    private static final Logger LOGGER = LoggerFactory.getLogger(ColumnsValidator.class);
    private final WorkbookService workbookService;

    public ColumnsValidator(WorkbookService workbookService) {
        this.workbookService = workbookService;
    }

    @Override
    public Boolean validate(String columnsStr, Map<ValidationResource, String> context) {
        if (columnsStr == null || columnsStr.trim().isEmpty()) {
            LOGGER.debug("Columns validation failed: value is null or empty");
            return false;
        }

        try {
            String driveId = context.get(ValidationResource.DRIVE_ID);
            String fileId = context.get(ValidationResource.FILE_ID);
            String tableName = context.get(ValidationResource.TABLE_NAME);
            
            List<String> columns = parseColumns(columnsStr);
            if (columns.isEmpty()) {
                LOGGER.debug("Columns validation failed: no columns provided");
                return false;
            }
            
            // Check if all columns exist in the table
            for (String column : columns) {
                if (!workbookService.isColumnExistsInTable(driveId, fileId, tableName, column)) {
                    LOGGER.debug("Columns validation failed: column '{}' does not exist in table '{}'", column, tableName);
                    return false;
                }
            }
            
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating columns: {}", cause.getMessage(), cause);
            return false;
        }
    }

    private List<String> parseColumns(String columnsStr) {
        // This is a simplified parsing - adjust based on actual format
        // Assuming columnsStr is a JSON array string or comma-separated list
        if (columnsStr.startsWith("[") && columnsStr.endsWith("]")) {
            columnsStr = columnsStr.substring(1, columnsStr.length() - 1);
        }
        return List.of(columnsStr.split(",\\s*"));
    }

    @Override
    public String getFetchFieldName() {
        return COLUMNS;
    }

    @Override
    public String getFieldType() {
        return "[ Text ]";
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter valid column names that exist in the table.";
    }

    @Override
    public String getConfirmStepMessage(String columns, Map<ValidationResource, String> context) {
        return "One or more columns do not exist in the specified table. Would you like to enter different column names?";
    }

    @Override
    public String getErrMessage(String columns) {
        return "Invalid column names. One or more columns do not exist in the specified table.";
    }
}