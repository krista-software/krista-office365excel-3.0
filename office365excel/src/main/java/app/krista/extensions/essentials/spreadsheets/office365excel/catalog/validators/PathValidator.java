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
import java.util.regex.Pattern;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.PATH;

/**
 * Validator for file paths in Office365Excel extension.
 * Ensures that paths are valid and properly formatted according to the expected format
 * used by Workbooks.getPath method.
 */
public class PathValidator implements Validator {
    private static final Logger LOGGER = LoggerFactory.getLogger(PathValidator.class);
    
    // Pattern to validate path format - allows folder paths with forward slashes
    // This matches paths like "/abc/cde" or "abc/cde" or just "abc" and allows spaces and more characters
    private static final Pattern VALID_PATH_PATTERN = Pattern.compile("^(/?([-\\w\\s]+)(/[-\\w\\s]+)*)?$");
    
    // Maximum path length to prevent excessively long paths
    private static final int MAX_PATH_LENGTH = 260;

    @Override
    public Boolean validate(String path, Map<ValidationResource, String> context) {
        if (path == null) {
            // Null paths are allowed and will be handled by emptyIfNull in Workbooks.getPath
            LOGGER.debug("Path validation: path is null, which is allowed");
            return true;
        }
        
        try {
            if (path.length() > MAX_PATH_LENGTH) {
                LOGGER.debug("Path validation failed: path exceeds maximum length of {} characters", MAX_PATH_LENGTH);
                return false;
            }
            
            if (!VALID_PATH_PATTERN.matcher(path).matches()) {
                LOGGER.debug("Path validation failed: path format is invalid: {}", path);
                return false;
            }
            
            if (path.contains("..") || path.contains("//") || path.contains("\\\\")) {
                LOGGER.debug("Path validation failed: path contains invalid sequences: {}", path);
                return false;
            }
            
            LOGGER.debug("Path validation passed for: {}", path);
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating path: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return PATH;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid folder path (e.g., '/Documents/Excel' or 'Documents/Excel').";
    }

    @Override
    public String getConfirmStepMessage(String path, Map<ValidationResource, String> context) {
        return "The path '" + path + "' is invalid. Would you like to enter a different path?";
    }

    @Override
    public String getErrMessage(String path) {
        return "Invalid folder path: " + path + ". Please provide a valid path format (e.g., '/Documents/Excel' or 'Documents/Excel').";
    }
}