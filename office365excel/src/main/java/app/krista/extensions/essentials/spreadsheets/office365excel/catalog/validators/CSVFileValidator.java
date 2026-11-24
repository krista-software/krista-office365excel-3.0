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
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Type;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.ROWS_FROM_CSV_FILE;

/**
 * Validator for CSV Files in Office365Excel extension.
 * <p>
 * This class validates that CSV files have valid file extensions by checking against
 * a predefined set of allowed file types. It provides user-facing messages for fetch,
 * confirmation, and error steps.
 * </p>
 */
public class CSVFileValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(CSVFileValidator.class);
    private static final Gson gson = new Gson();

    private final static Set<String> validFileExtensions = Set.of(
            // CSV and related formats
            "csv"
    );

    @Override
    public Boolean validate(String csvFilesStr, Map<ValidationResource, String> context) {
        try {
            if (csvFilesStr == null || csvFilesStr.trim().isEmpty()) {
                LOGGER.debug("CSV file validation: No files provided");
                return false;
            }

            // Try to parse as JSON array of file objects
            Type listType = new TypeToken<List<Map<String, Object>>>() {}.getType();
            List<Map<String, Object>> files = gson.fromJson(csvFilesStr, listType);

            if (files == null || files.isEmpty()) {
                // Empty list is valid if files are optional
                LOGGER.debug("CSV file validation: Empty file list");
                return false;
            }

            // Check if all files have valid file extensions
            for (Map<String, Object> file : files) {
                if (file == null) {
                    continue;
                }

                // Get the file name from the file object
                String fileName = null;
                if (file.containsKey("name")) {
                    fileName = (String) file.get("name");
                } else if (file.containsKey("fileName")) {
                    fileName = (String) file.get("fileName");
                } else if (file.containsKey("originalName")) {
                    fileName = (String) file.get("originalName");
                }

                if (fileName == null || fileName.trim().isEmpty()) {
                    LOGGER.debug("CSV file validation failed: No file name found");
                    return false; // No file name found
                }

                // Check if the file extension is valid
                String lowerCaseFileName = fileName.toLowerCase();
                String fileExtension = getFileExtension(lowerCaseFileName);
                if (fileExtension == null || !validFileExtensions.contains(fileExtension)) {
                    LOGGER.debug("CSV file validation failed: Invalid file extension: {}", fileExtension);
                    return false; // Invalid file extension found
                }
            }

            LOGGER.debug("CSV file validation passed for all files");
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating CSV files: {}", cause.getMessage(), cause);
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
        return ROWS_FROM_CSV_FILE;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.FILE_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please upload a valid data file. Supported format: CSV.";
    }

    @Override
    public String getConfirmStepMessage(String csvFilesStr, Map<ValidationResource, String> context) {
        return "The provided file is not a valid data file. Please upload a file in CSV format.";
    }

    @Override
    public String getErrMessage(String csvFilesStr) {
        return "Invalid file format. Only CSV files are allowed.";
    }
}
