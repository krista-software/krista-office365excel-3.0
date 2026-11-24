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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.TABLE_NAME;

/**
 * Validator for Table Names in Office365Excel extension.
 */
public class TableNameValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(TableNameValidator.class);
    private final WorkbookService workbookService;

    public TableNameValidator(WorkbookService workbookService) {
        this.workbookService = workbookService;
    }

    @Override
    public Boolean validate(String tableName, Map<ValidationResource, String> context) {
        if (tableName == null || tableName.trim().isEmpty()) {
            return false;
        }
        
        try {
            String driveId = context.get(ValidationResource.DRIVE_ID);
            String fileId = context.get(ValidationResource.FILE_ID);
            
            return workbookService.isTableAccessible(driveId, fileId, tableName);
        } catch (Exception cause) {
            LOGGER.error("Error validating table name: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return TABLE_NAME;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter the name of an existing Excel table. Table names are case-sensitive and can be found in the Excel file's table list.";
    }

    @Override
    public String getConfirmStepMessage(String tableName, Map<ValidationResource, String> context) {
        return "We couldn't find the table '" + tableName + "' in this workbook. Please check that the table name is spelled correctly and exists in the file. Would you like to try a different table name?";
    }

    @Override
    public String getErrMessage(String tableName) {
        return "Table '" + tableName + "' not found or inaccessible. Please verify the table exists in the workbook and you have permission to access it.";
    }
}