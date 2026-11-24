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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonElement;
import com.microsoft.graph.models.WorkbookTable;

import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public final class Tables {
    private Tables() {
    }

    public static List<FreeForm> getRowsDataForColumns(List<FreeForm> rowDataList, String keyColumnHeader,
                                                       List<String> keyRows, List<String> lookupColumns,
                                                       Boolean getAllRecord) {
        List<FreeForm> allAvailableData = new ArrayList<>();

        if (rowDataList != null && !rowDataList.isEmpty()
                && keyRows != null && !keyRows.isEmpty()
                && lookupColumns != null && !lookupColumns.isEmpty()) {

            lookupColumns = lookupColumns.stream().map(String::toLowerCase).collect(Collectors.toList());

            keyRows = keyRows.stream().map(keyRow -> keyRow.trim().toLowerCase()).collect(Collectors.toList());
            for (String keyRow : keyRows) {
                String keyRowValue = keyRow.toLowerCase().trim(); //hari

                // Name: hari, dob: 10/1/1996, standard: 1
                for (FreeForm rowData : rowDataList) {
                    String keyColumnValue = rowData.get(keyColumnHeader).toString();

                    if (keyColumnValue.toLowerCase().equals(keyRowValue)) // hari
                    {
                        if (getAllRecord)
                            allAvailableData.add(getFreeFormRecord(keyColumnHeader, lookupColumns, rowData, keyColumnValue));
                        else {
                            if (!isRecordExist(allAvailableData, keyColumnHeader, keyRowValue))
                                allAvailableData.add(getFreeFormRecord(keyColumnHeader, lookupColumns, rowData, keyColumnValue));
                        }
                    }
                }
            }
            return allAvailableData;
        }
        return null;
    }

    private static FreeForm getFreeFormRecord(String keyColumnHeader, List<String> lookupColumns, FreeForm rowData,
                                              String keyColumnValue) {
        FreeForm freeForm = new FreeForm();
        try {
            if (!keyColumnValue.isEmpty() && Numbers.isNumber(keyColumnValue)) {
                freeForm.put(keyColumnHeader, "Number", Double.parseDouble(keyColumnValue));
            } else {
                freeForm.put(keyColumnHeader, "Text", keyColumnValue);
            }
        } catch (NumberFormatException exception) {
            freeForm.put(keyColumnHeader, "Text", keyColumnValue);
        }

        if (!rowData.isEmpty()) {
            Set<String> setOfKeys = rowData.keySet();
            for (String columnHeaderKey : setOfKeys) {
                if (lookupColumns.contains(columnHeaderKey.toLowerCase().trim())) {
                    try {
                        if (!String.valueOf(rowData.get(columnHeaderKey)).isEmpty() && Numbers.isNumber(String.valueOf(rowData.get(columnHeaderKey).toString()))) {
                            freeForm.put(columnHeaderKey, "Number", Double.parseDouble(rowData.get(columnHeaderKey).toString()));
                        } else {
                            freeForm.put(columnHeaderKey, "Text", rowData.get(columnHeaderKey));
                        }
                    } catch (NumberFormatException numberFormatException) {
                        freeForm.put(columnHeaderKey, "Text", rowData.get(columnHeaderKey));
                    }

                }
            }
        }
        return freeForm;
    }

    private static boolean isRecordExist(List<FreeForm> allAvailableData, String keyColumnHeader, String keyRowValue) {

        if (allAvailableData == null || allAvailableData.isEmpty()) {
            return false;
        }

        for (FreeForm availableData : allAvailableData) {
            String value = availableData.get(keyColumnHeader).toString();
            if (value.trim().equalsIgnoreCase(keyRowValue)) {
                return true;
            }
        }

        return false;
    }

    private static List<String> addEmptyString(List<String> eachRow, Integer noOfColumns) {
        List<String> modifiableRow = new ArrayList<>(eachRow); // ensure modifiable
        int emptyStringToBeAdded = noOfColumns - modifiableRow.size();
        for (int i = 0; i < emptyStringToBeAdded; i++) {
            modifiableRow.add(Constants.EMPTY_STRING);
        }
        return modifiableRow;
    }

    private static List<String> getOnlyRequiredRowData(List<String> eachRow, Integer noOfColumns) {
        List<String> requiredRowData = new ArrayList<>();
        if (eachRow.size() == noOfColumns) {
            return eachRow;
        }
        for (int i = 0; i < noOfColumns; i++) {
            requiredRowData.add(eachRow.get(i));
        }
        return requiredRowData;
    }


    public static List<Table> getTables(List<WorkbookTable> workbookTables, String fileName) {
        List<Table> tableList = new ArrayList<>();
        if (workbookTables.isEmpty() || fileName.isEmpty()) {
            return List.of();
        }
        for (WorkbookTable workbookTable : workbookTables) {
            Table excelSheet = getTable(workbookTable, fileName);
            tableList.add(excelSheet);
        }

        return tableList;
    }

    public static Table getTable(WorkbookTable workbookTable, String fileName) {
        return new Table(workbookTable.id, workbookTable.name, fileName);
    }

    public static JsonElement transformRows(List<String> rows, Integer noOfColumns, String separator) {
        List<List<Object>> values = new ArrayList<>();
        if (separator == null || separator.isEmpty()) {
            separator = Constants.COMMA;
        }
        for (String row : rows) {
            List<String> rowData = new ArrayList<>(splitMe(row, separator));
            if (rowData.size() < noOfColumns) {
                rowData = addEmptyString(rowData, noOfColumns);
            } else {
                rowData = getOnlyRequiredRowData(rowData, noOfColumns);
            }

            List<Object> transformedRowData = new ArrayList<>();
            for (String data : rowData) {
                transformedRowData.add(transformData(data));
            }
            values.add(transformedRowData);
        }
        return Constants.GSON.toJsonTree(values);
    }

    private static Object transformData(String data) {
        try {
            return Double.parseDouble(data);
        } catch (NumberFormatException exception) {
            return data;
        }
    }

    /**
     * Splits the given row into a list of strings based on the specified separator.
     * The method uses the entire provided separator string to split the row.
     *
     * @param row             the string to be split, typically representing a row of data
     * @param separatorString the string that acts as the delimiter for splitting the row.
     *                        This can be more than one character and is used directly as the separator.
     * @return a list of strings resulting from splitting the row by the separator
     * @throws NullPointerException if the row or separatorString is null
     */
    private static List<String> splitMe(String row, String separatorString) {
        separatorString = Pattern.quote(separatorString);
        String[] split = row.split(separatorString);
        return Arrays.asList(split);
    }

}

