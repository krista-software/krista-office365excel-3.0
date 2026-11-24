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

import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.google.gson.reflect.TypeToken;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Type;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes.FILE_FIELD_LOGICAL_NAME;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.EXCEL_FILE;

/**
 * Validator for Excel Files in Office365Excel extension.
 * <p>
 * This class validates that files have valid Excel file extensions by checking against
 * a predefined set of allowed file types. It provides user-facing messages for fetch,
 * confirmation, and error steps.
 * </p>
 */
public class ExcelFileValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelFileValidator.class);

    private final static Set<String> validFileExtensions = Set.of(
            // Excel formats
            "xlsx", "xlsm", "xls", "xlsb"
    );

    @Override
    public Boolean validate(String excelFilesStr, Map<ValidationResource, String> context) {
        try {
            if (excelFilesStr == null || excelFilesStr.trim().isEmpty()) {
                LOGGER.debug("Excel file validation: No files provided");
                return false;
            }

            // Try to parse as JSON array of file objects
            Type listType = new TypeToken<List<Map<String, Object>>>() {}.getType();
            List<Map<String, Object>> files = Constants.GSON.fromJson(excelFilesStr, listType);

            if (files == null || files.isEmpty()) {
                LOGGER.debug("Excel file validation: Empty file list");
                return false;
            }

            for (Map<String, Object> file : files) {
                // Extract file name
                String fileName = (String) file.get("name");
                if (fileName == null || fileName.trim().isEmpty()) {
                    LOGGER.debug("Excel file validation failed: Missing file name");
                    return false;
                }

                // Check if the file extension is valid
                String lowerCaseFileName = fileName.toLowerCase();
                String fileExtension = getFileExtension(lowerCaseFileName);
                if (fileExtension == null || !validFileExtensions.contains(fileExtension)) {
                    LOGGER.debug("Excel file validation failed: Invalid file extension: {}", fileExtension);
                    return false; // Invalid file extension found
                }
            }

            LOGGER.debug("Excel file validation passed for all files");
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating Excel files: {}", cause.getMessage(), cause);
            return false;
        }
    }

    /**
     * Extracts the file extension from a filename.
     *
     * @param fileName the filename (should be lowercase)
     * @return the file extension without the dot, or null if no extension found
     */
    private String getFileExtension(String fileName) {
        if (fileName == null || fileName.trim().isEmpty()) {
            return null;
        }

        int lastDotIndex = fileName.lastIndexOf('.');
        if (lastDotIndex == -1 || lastDotIndex == fileName.length() - 1) {
            return null; // No extension or dot at the end
        }

        return fileName.substring(lastDotIndex + 1);
    }

    @Override
    public String getFetchFieldName() {
        return EXCEL_FILE;
    }

    @Override
    public String getFieldType() {
        return FILE_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please upload a valid Excel file. Supported formats: .xlsx, .xlsm, .xls, .xlsb.";
    }

    @Override
    public String getConfirmStepMessage(String excelFilesStr, Map<ValidationResource, String> context) {
        return "The provided file is not a valid Excel file. Please upload a file in one of the supported formats (.xlsx, .xlsm, .xls, .xlsb).";
    }

    @Override
    public String getErrMessage(String excelFilesStr) {
        return "Invalid file format. Only Excel files (.xlsx, .xlsm, .xls, .xlsb) are allowed.";
    }
}