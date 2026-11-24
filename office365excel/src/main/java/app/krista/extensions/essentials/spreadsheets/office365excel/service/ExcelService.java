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

package app.krista.extensions.essentials.spreadsheets.office365excel.service;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.ItemReference;
import com.microsoft.graph.models.RemoteItem;
import org.jetbrains.annotations.NotNull;

import java.util.ArrayList;
import java.util.List;

public class ExcelService {

    public ExcelService() {
    }

    @NotNull
    public static List<ExcelFile> getExcelFiles(String fileName, List<DriveItem> currentDriveItems) {
        List<ExcelFile> allSharedExcelFiles = new ArrayList<>();
        currentDriveItems.stream()
                .filter(driveItem -> driveItem.name != null && StringUtil.isExcelFile(driveItem.name))
                .forEachOrdered(driveItem -> {
                    RemoteItem remoteItem = driveItem.remoteItem;
                    String driveId;
                    if (remoteItem != null && remoteItem.parentReference != null) {
                        driveId = remoteItem.parentReference.driveId;
                    } else {
                        // SharePoint items may not have remoteItem; fall back to the item's own parentReference
                        ItemReference parentRef = driveItem.parentReference;
                        driveId = parentRef != null ? parentRef.driveId : null;
                    }
                    ExcelFile excelFile = new ExcelFile(driveItem.id,
                            driveId != null ? driveId : "", driveItem.name, driveItem.webUrl);
                    if (fileName == null) {
                        // Add all files
                        allSharedExcelFiles.add(excelFile);
                    } else if (driveItem.name.equalsIgnoreCase(fileName)) {
                        // Add only file name matching files
                        allSharedExcelFiles.add(excelFile);
                    }
                });
        return allSharedExcelFiles;
    }

}
