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
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import com.microsoft.graph.http.GraphServiceException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.Map;
import java.util.regex.Pattern;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.SHEET_NAME;

/**
 * Validator for Excel sheet names in Office365Excel extension.
 * <p>
 * This class validates that sheet names follow Excel's naming conventions:
 * - Cannot be empty
 * - Cannot exceed 31 characters
 * - Cannot contain certain special characters: [ ] * ? / \
 * - Cannot start or end with a single quote
 * </p>
 */
public class SheetNameValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(SheetNameValidator.class);
    private static final Pattern INVALID_CHARS_PATTERN = Pattern.compile("[\\[\\]\\*\\?\\/\\\\]");

    private final WorksheetService worksheetService;

    /**
     * Creates a validator that can optionally check if the sheet exists.
     *
     * @param worksheetService The service to use for checking sheet existence
     */
    @Inject
    public SheetNameValidator(WorksheetService worksheetService) {
        this.worksheetService = worksheetService;
    }

    @Override
    public Boolean validate(String sheetName, Map<ValidationResource, String> context) {
        if (sheetName == null || sheetName.trim().isEmpty()) {
            LOGGER.info("Sheet name validation failed: name is null or empty");
            return false;
        }

        try {
            // Check for invalid characters
            if (INVALID_CHARS_PATTERN.matcher(sheetName).find()) {
                LOGGER.info("Sheet name validation failed: name contains invalid characters");
                return false;
            }

            // Check for starting or ending with a single quote
            if (sheetName.startsWith("'") || sheetName.endsWith("'")) {
                LOGGER.info("Sheet name validation failed: name starts or ends with a single quote");
                return false;
            }

            // If worksheetService is provided, verify the sheet exists
            if (worksheetService != null) {
                String fileId = context.get(ValidationResource.FILE_ID);
                String driveId = context.get(ValidationResource.DRIVE_ID);
                if (fileId == null || fileId.trim().isEmpty()) {
                    LOGGER.info("Sheet existence check failed: file ID is missing");
                    return false;
                }

                LOGGER.info("Validating sheet existence with fileId='{}', driveId='{}', sheetName='{}'", fileId, driveId, sheetName);
                boolean exists = worksheetService.worksheetExists(fileId, sheetName, driveId);
                if (!exists) {
                    LOGGER.info("Sheet existence check failed: sheet '{}' does not exist in file {} (drive: {})", sheetName, fileId, driveId);
                    return false;
                }
            }

            LOGGER.info("Sheet name validation passed for: {}", sheetName);
            return true;
        } catch (GraphServiceException cause) {
            LOGGER.error("Graph API error validating sheet name: {}", cause.getMessage(), cause);
            throw cause;
        } catch (IllegalStateException cause) {
            // Access denied (403) or other non-recoverable error from worksheetExists;
            // must propagate so the caller sees the real error, not "sheet not found".
            LOGGER.error("Access error validating sheet name: {}", cause.getMessage(), cause);
            throw cause;
        } catch (Exception cause) {
            LOGGER.error("Error validating sheet name: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return SHEET_NAME;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        if (worksheetService != null) {
            return "Please enter the name of an existing Excel sheet. You can find sheet names by using the 'Get Sheets Information' request.";
        }
        return "Please enter a valid Excel sheet name. Sheet names must be 31 characters or less and cannot contain these special characters: [ ] * ? / \\";
    }

    @Override
    public String getConfirmStepMessage(String sheetName, Map<ValidationResource, String> context) {
        if (worksheetService != null) {
            return "The sheet '" + sheetName + "' doesn't exist in this workbook. Please check the sheet name for typos or use 'Get Sheets Information' to see available sheets. Would you like to try a different sheet name?";
        }
        return "The sheet name '" + sheetName + "' is not valid. Excel sheet names must be 31 characters or less and cannot contain these characters: [ ] * ? / \\. Would you like to try again?";
    }

    @Override
    public String getErrMessage(String sheetName) {
        if (worksheetService != null) {
            return "Sheet '" + sheetName + "' not found in this workbook. Please verify the sheet name is correct and the sheet exists.";
        }
        return "Invalid sheet name '" + sheetName + "'. Excel sheet names must be 31 characters or less and cannot contain these characters: [ ] * ? / \\";
    }
}