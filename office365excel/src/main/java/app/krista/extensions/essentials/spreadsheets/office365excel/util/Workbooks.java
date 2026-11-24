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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.WorkbookFunctionsVlookupParameterSet;
import com.microsoft.graph.models.WorkbookWorksheetRangeParameterSet;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil.emptyIfNull;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil.isNullOrEmpty;

/**
 * Helper class which provides helper methods for workbook operations.
 */
public final class Workbooks {

    private Workbooks() {

    }

    /**
     * @return String that denotes "workbook full path".
     */
    public static String getPath(String workbookFolderPath, String workbookName) {
        return emptyIfNull(workbookFolderPath) + Constants.BACKWARD_SLASH + validateWorkbookName(workbookName);
    }

    /**
     * @return null/workbookName.
     */
    public static String validateWorkbookName(String workbookName) {
        if (isNullOrEmpty(workbookName)) {
            return null;
        }
        if (!workbookName.endsWith(Constants.EXCEL_FILE_EXTENSION)) {
            return workbookName + Constants.EXCEL_FILE_EXTENSION;
        }
        return workbookName;
    }

    public static List<ExcelFile> getWorkbooks(String workbookName, List<DriveItem> driveItems) {
        List<ExcelFile> workbooks = new ArrayList<>();
        for (DriveItem driveItem : driveItems) {
            if (driveItem.name != null && driveItem.name.toLowerCase().contains(workbookName.toLowerCase())
                    && driveItem.name.endsWith(Constants.EXCEL_FILE_EXTENSION) && driveItem.parentReference != null) {
                workbooks.add(new ExcelFile(driveItem.id, driveItem.parentReference.driveId, driveItem.name, driveItem.webUrl));
            }
        }
        return workbooks;
    }

    public static FreeForm getFreeForm(String lookupValue, Object value) {
        FreeForm freeForm = new FreeForm();
        freeForm.put(lookupValue, Constants.TEXT, value == null ? "" : value);
        return freeForm;
    }

    public static WorkbookFunctionsVlookupParameterSet getvLookupParameterSet(JsonElement lookupVal, JsonElement colIndex, JsonElement rangeLookup, JsonElement rangeJsonElem) {
        return WorkbookFunctionsVlookupParameterSet
                .newBuilder()
                .withLookupValue(lookupVal)
                .withRangeLookup(rangeLookup)
                .withColIndexNum(colIndex)
                .withTableArray(rangeJsonElem)
                .build();
    }

    public static WorkbookWorksheetRangeParameterSet getWorkbookWorksheetRangeParameterSet(String range) {
        return WorkbookWorksheetRangeParameterSet
                .newBuilder()
                .withAddress(range)
                .build();
    }

    public static List<FreeForm> getFreeForms(List<FreeForm> finalData, List<FreeForm> newData) {
        List<FreeForm> newDataModifiedList = new ArrayList<>();
        for (FreeForm freeForm : newData) {
            for (String key : freeForm.keySet()) {
                if (!finalData.isEmpty() && finalData.getFirst() != null &&
                        finalData.getFirst().getTypeDescription(key)
                                != null && !finalData.getFirst().getTypeDescription(key).equals(freeForm.getTypeDescription(key))) {
                    Object value = freeForm.get(key);
                    if (finalData.getFirst().getTypeDescription(key).equals(Constants.TEXT) &&
                            freeForm.getTypeDescription(key).equals("Number")) {
                        DecimalFormat format = new DecimalFormat("0.#");
                        value = format.format(value);
                    }
                    freeForm.put(key, finalData.getFirst().getTypeDescription(key), value);
                }
            }
            newDataModifiedList.add(freeForm);
        }
        return newDataModifiedList;
    }

    public static void extractRequiredData(FreeForm finalRow, List<String> dummyRows, JsonArray jsonArray) {
        for (int i = 0; i < jsonArray.size(); i++) {
            String s = jsonArray.get(i).getAsJsonPrimitive().getAsString();
            dummyRows.add(s);
        }
        StringUtil.removeEmptyStringsFromEnd(dummyRows);
        for (int j = 0; j < dummyRows.size(); j++) {
            finalRow.put(generateColumnLabel(j + 1), Constants.TEXT, dummyRows.get(j));
        }
    }

    private static String generateColumnLabel(int columnNumber) {
        if (columnNumber <= 0) {
            throw new IllegalArgumentException("Column number must be greater than zero.");
        }

        StringBuilder columnLabel = new StringBuilder();

        while (columnNumber > 0) {
            columnNumber--;
            columnLabel.insert(0, (char) ('A' + columnNumber % 26));
            columnNumber /= 26;
        }
        return columnLabel.toString();
    }

    public static String extractRange(String input) {

        String[] parts = input.split("!");
        if (parts.length != 2) {
            throw new IllegalArgumentException(
                    "Invalid input format. Expected format: SheetName!A1:C1");
        }

        return parts[1];
    }
}
