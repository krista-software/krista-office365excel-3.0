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

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.DRIVE_ID;

/**
 * Validator for Drive IDs in Office365Excel extension.
 */
public class DriveIdValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(DriveIdValidator.class);
    private final WorkbookService workbookService;

    public DriveIdValidator(WorkbookService workbookService) {
        this.workbookService = workbookService;
    }

    @Override
    public Boolean validate(String driveId, Map<ValidationResource, String> context) {
        if (driveId == null || driveId.trim().isEmpty()) {
            return false;
        }
        
        try {
            return workbookService.isDriveAccessible(driveId);
        } catch (Exception cause) {
            LOGGER.error("Error validating drive ID: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return DRIVE_ID;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a valid OneDrive or SharePoint drive ID. You can find this using the 'Get All Shared Drives Information' request.";
    }

    @Override
    public String getConfirmStepMessage(String driveId, Map<ValidationResource, String> context) {
        return "We couldn't access the drive with ID '" + driveId + "'. This might be because the drive doesn't exist or you don't have permission to access it. Would you like to try a different Drive ID?";
    }

    @Override
    public String getErrMessage(String driveId) {
        return "Unable to access drive with ID '" + driveId + "'. Please verify the drive exists and you have the necessary permissions.";
    }
}