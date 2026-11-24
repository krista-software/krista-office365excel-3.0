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

package app.krista.extensions.essentials.spreadsheets.office365excel.util;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Drive;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.RemoteItem;
import org.jetbrains.annotations.NotNull;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

import static app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService.TEXT;

public class Result {

    private Result() {

    }

    private static final Logger LOGGER = LoggerFactory.getLogger(Result.class);

    @NotNull
    public static List<Drive> getDrives(List<DriveItem> currentDriveItems) {
        Set<String> seenDriveIds = new HashSet<>();
        List<Drive> drives = new ArrayList<>();
        for (DriveItem driveItem : currentDriveItems) {
            RemoteItem remoteItem = driveItem.remoteItem;
            if (remoteItem == null || remoteItem.parentReference == null) {
                continue;
            }
            String driveId = remoteItem.parentReference.driveId;
            if (driveId == null || driveId.isBlank() || !seenDriveIds.add(driveId)) {
                continue;
            }
            String sharedBy = null;
            if (remoteItem.shared != null && remoteItem.shared.sharedBy != null) {
                if (remoteItem.shared.sharedBy.user != null) {
                    sharedBy = remoteItem.shared.sharedBy.user.displayName;
                } else if (remoteItem.shared.sharedBy.application != null) {
                    sharedBy = remoteItem.shared.sharedBy.application.displayName;
                }
            }
            if (sharedBy == null) {
                sharedBy = "SharePoint";
            }
            LOGGER.info("getDrives: driveId={} sharedBy={}", driveId, sharedBy);
            drives.add(new Drive(driveId, sharedBy, null));
        }
        return drives;
    }

    @NotNull
    public static Map<String, Object> getResultFromMinMaxValueInformation(List<Object> keyColumnData, Map<String, Double> minMaxValueInformation) {

        double maxValueRowIndex = minMaxValueInformation.get(Constants.MAX_VALUE_ROW_INDEX);
        double minValueRowIndex = minMaxValueInformation.get(Constants.MIN_VALUE_ROW_INDEX);

        Object keyMaxColumnValue = keyColumnData.get((int) maxValueRowIndex);
        Object keyMinColumnValue = keyColumnData.get((int) minValueRowIndex);
        String maxColumnValueType = getValueType(keyMaxColumnValue);
        String minColumnValueType = getValueType(keyMinColumnValue);
        FreeForm freeForm = new FreeForm();
        freeForm.put("Max Value", "Number", minMaxValueInformation.get(Constants.MAX_VALUE));
        freeForm.put("Max Value Key", maxColumnValueType, keyMaxColumnValue);
        freeForm.put("Min Value", "Number", minMaxValueInformation.get(Constants.MIN_VALUE));
        freeForm.put("Min Value Key", minColumnValueType, keyMinColumnValue);
        LOGGER.info("minMaxValues:{}", freeForm);

        return freeForm;
    }

    private static String getValueType(Object value) {
        if (value instanceof Number) {
            return "Number";
        }
        if (value instanceof Date) {
            return "Date";
        }
        return TEXT;
    }
}
