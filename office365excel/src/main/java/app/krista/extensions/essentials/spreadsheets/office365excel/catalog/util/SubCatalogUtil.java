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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class SubCatalogUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(SubCatalogUtil.class);

    /**
     * Extracts files from a map object that can contain either a single File or a List of Files
     * @param fileObj The object from the map that may contain File(s)
     * @return List of File objects, empty list if no valid files found
     */
    public static List<File> extractFilesFromMap(Object fileObj) {
        List<File> rowsFromFileList = new ArrayList<>();

        if (fileObj instanceof List<?> files) {
            for (Object item : files) {
                if (item instanceof File) {
                    rowsFromFileList.add((File) item);
                }
            }
        } else if (fileObj instanceof File) {
            rowsFromFileList.add((File) fileObj);
        } else {
            LOGGER.debug("No valid files found in the input map");
        }

        return rowsFromFileList;
    }

    /**
     * Processes files from input map and updates state map with the appropriate file list
     * @param inputMap The input map containing file data
     * @param stateMap The state map to be updated
     * @param state The existing state containing fallback file data
     * @param fileKey The key used to retrieve files from input map and store in state map
     */
    public static void processAndUpdateFiles(Map<String, Object> inputMap, Map<String, Object> stateMap,
                                           Map<String, Object> state, String fileKey) {
        Object fileObj = inputMap.get(fileKey);
        List<File> rowsFromFileList = extractFilesFromMap(fileObj);

        if (!rowsFromFileList.isEmpty()) {
            stateMap.put(fileKey, rowsFromFileList);
        } else {
            stateMap.put(fileKey, state.get(fileKey));
        }
    }
}
