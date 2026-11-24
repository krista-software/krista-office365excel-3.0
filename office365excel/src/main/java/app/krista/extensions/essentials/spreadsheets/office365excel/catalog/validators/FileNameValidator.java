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
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.SHARED_FILE_NAME;

/**
 * Validator for shared Excel file names
 */
public class FileNameValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(FileNameValidator.class);

    @Override
    public Boolean validate(String fileName, Map<ValidationResource, String> context) {
        if (fileName == null || fileName.trim().isEmpty()) {
            LOGGER.debug("Shared file name validation failed: name is empty");
            return false;
        }
        
        // Check if the file has an Excel extension
        if (!StringUtil.isExcelFile(fileName)) {
            LOGGER.debug("Shared file name validation failed: not an Excel file: {}", fileName);
            return false;
        }
        
        LOGGER.debug("Shared file name validation passed for: {}", fileName);
        return true;
    }

    @Override
    public String getFetchFieldName() {
        return SHARED_FILE_NAME;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid Excel file name including the extension. Supported formats: .xlsx (Excel Workbook), .xlsm (Excel Macro-Enabled), .xls (Excel 97-2003), or .xlsb (Excel Binary).";
    }

    @Override
    public String getConfirmStepMessage(String fileName, Map<ValidationResource, String> context) {
        return "The file '" + fileName + "' is not recognized as an Excel file. Please ensure the file name ends with a valid Excel extension (.xlsx, .xlsm, .xls, or .xlsb). Would you like to try again?";
    }

    @Override
    public String getErrMessage(String fileName) {
        return "Invalid Excel file name '" + fileName + "'. Please include a valid Excel extension: .xlsx, .xlsm, .xls, or .xlsb.";
    }
}