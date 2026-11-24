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

package app.krista.extensions.essentials.spreadsheets.office365excel.util.parser;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Table;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Numbers;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.FreeForm;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonPrimitive;
import com.microsoft.graph.models.WorkbookTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

public final class DataParser {

    public static final String NUMBER = "Number";
    public static final String TEXT = "Text";
    private static final Logger LOGGER = LoggerFactory.getLogger(DataParser.class);


    private DataParser() {
    }

    private static final Map<String, String> headerDataType = new LinkedHashMap<>();

    public static List<FreeForm> parse(JsonElement element) {
        List<FreeForm> columns = new ArrayList<>();
        JsonArray rows = element.getAsJsonArray();
        if (rows == null) {
            throw new IllegalArgumentException("Invalid rows passed to the function");
        }
        //if the sheet is blank and if you try to add the rows than it would add blank line in first row
        if (rows.get(0).toString().equals("[\"\"]")) {
            return Collections.emptyList();
        }
        FreeForm previousRow = null;
        for (JsonElement row : rows) {
            JsonArray rowArray = row.getAsJsonArray();
            FreeForm column = getColumn(rowArray, previousRow);
            previousRow = column;
            columns.add(column);
        }
        return columns;
    }

    private static FreeForm getColumn(JsonArray row, FreeForm previousRow) {
        FreeForm column = new FreeForm();
        for (int index = 0; index < row.size(); index++) {
            JsonPrimitive rowDataAtIndex = row.get(index).getAsJsonPrimitive();
            if (rowDataAtIndex == null) {
                throw new IllegalArgumentException("rowData are not valid");
            }
            try {
                LOGGER.debug("element {} : {}", (index + 1), rowDataAtIndex.getAsString());
                if (rowDataAtIndex.isNumber() && (previousRow == null || previousRow.getTypeDescription(getHeader(index)).equals(
                        NUMBER))) {
                    column.put(getHeader(index), NUMBER, Double.parseDouble(rowDataAtIndex.getAsString()));
                } else {
                    column.put(getHeader(index), TEXT, rowDataAtIndex.getAsString());
                }
            } catch (NumberFormatException exception) {
                LOGGER.error("exception occured: {}", rowDataAtIndex.getAsString());
                column.put(getHeader(index), TEXT, rowDataAtIndex.toString());
            }
        }
        return column;
    }

    private static FreeForm populateColumnValue(List<String> headers, JsonArray rowArray) {
        FreeForm rows = new FreeForm();
        for (int index = 0; index < rowArray.size(); index++) {
            JsonPrimitive rowDataAtIndex = rowArray.get(index).getAsJsonPrimitive();
            Object value = rowDataAtIndex.getAsString();
            rows.put(headers.get(index), TEXT, value);
        }
        return rows;
    }

    private static FreeForm populateColumnValueForTableData(List<String> headers, JsonArray rowArray) {
        FreeForm rows = new FreeForm();

        for (int index = 0; index < rowArray.size(); index++) {
            JsonPrimitive rowDataAtIndex = rowArray.get(index).getAsJsonPrimitive();
            String rowDataStr = rowDataAtIndex.getAsString();
            try {
                if (!rowDataStr.isEmpty() && Numbers.isNumber(rowDataStr)) {
                    rows.put(headers.get(index), NUMBER, Double.valueOf(rowDataStr));
                } else {
                    rows.put(headers.get(index), TEXT, rowDataStr);
                }
            } catch (NumberFormatException numberFormatException) {
                LOGGER.error("exception occured:{}", rowDataStr);
                rows.put(headers.get(index), TEXT, rowDataStr);
            }
        }

        return rows;
    }

    private static List<FreeForm> fillTheTableData(List<FreeForm> insufficientRows, List<String> headers) {
        for (FreeForm freeForm : insufficientRows) {
            if (freeForm.size() < headers.size()) {
                List<String> missingKeys = headers.stream().filter(e -> !freeForm.containsKey(e)).toList();
                for (String missingKey : missingKeys) {
                    String datatype = headerDataType.get(missingKey);
                    if (datatype != null && datatype.equals(TEXT)) {
                        datatype = TEXT;
                    } else {
                        datatype = NUMBER;
                    }
                    Object value = Objects.equals(datatype, TEXT) ? "" : 0.0;
                    freeForm.put(missingKey, datatype, value);
                }

            }
        }
        return insufficientRows;
    }

    public static List<FreeForm> parseTableData(JsonElement jsonElement) {
        List<String> headers = getColumnHeaders(jsonElement);
        JsonArray rows = jsonElement.getAsJsonArray();
        List<FreeForm> rowsResult = new ArrayList<>();
        //starting from 1 as we are already getting the header as first row so need to start from the next
        for (int index = 1; index < rows.size(); index++) {
            JsonArray rowArray = rows.get(index).getAsJsonArray();

            for (JsonElement element : rowArray) {
                if (element != null && !element.getAsString().isBlank()) {
                    FreeForm row = populateColumnValueForTableData(headers, rowArray);
                    rowsResult.add(row);
                    break;
                }
            }
        }
        return fillTheTableData(rowsResult, headers);
    }


    public static List<FreeForm> parseTableDataForColumnValue(JsonElement jsonElement) {
        List<String> headers = getColumnHeaders(jsonElement);
        JsonArray rows = jsonElement.isJsonObject() ? jsonElement.getAsJsonObject().get("text").getAsJsonArray() : jsonElement.getAsJsonArray();
        List<FreeForm> rowsResult = new ArrayList<>();
        //starting from 1 as we are already getting the header as first row so need to start from the next
        for (int index = 1; index < rows.size(); index++) {
            JsonArray rowArray = rows.get(index).getAsJsonArray();

            for (JsonElement element : rowArray) {
                if (element != null && !element.getAsString().isBlank()) {
                    FreeForm row = populateColumnValue(headers, rowArray);
                    rowsResult.add(row);
                    break;
                }
            }
        }
        return fillTheTableData(rowsResult, headers);
    }

    /**
     * Offset represents the distance between "A" and targetCharacter.
     * Example: distance between A and C will be 2, so offset will be 2.
     *
     * @return String i.e. targetCharacter.
     */
    private static String getHeader(int offSet) {
        Character baseCharacter = Constants.A;
        int division = offSet / 26;
        StringBuilder header = new StringBuilder(Constants.EMPTY_STRING);
        while (division != 0) {
            header.append(baseCharacter);
            division--;
        }
        char character = (char) (baseCharacter + offSet % 26);
        return header.toString() + character;
    }

    public static List<String> getColumnHeaders(JsonElement jsonElement) {
        List<String> headers = new ArrayList<>();
        if (jsonElement == null) return headers;
        //key text would always be present in JsonObject
        JsonArray rows = jsonElement.isJsonObject() ? jsonElement.getAsJsonObject().get("text").getAsJsonArray() : jsonElement.getAsJsonArray();
        JsonArray rowArray = rows.get(0).getAsJsonArray();
        for (int index = 0; index < rowArray.size(); index++) {
            JsonPrimitive rowDataAtIndex = rowArray.get(index).getAsJsonPrimitive();
            String key = rowDataAtIndex.getAsString();
            headers.add(key);

        }
        return headers;
    }

    public static Object getCellData(JsonElement values) {
        if (values.isJsonArray()) {
            return getData(values);
        }
        if (values.isJsonPrimitive()) {
            JsonPrimitive jsonPrimitive = values.getAsJsonPrimitive();
            return jsonPrimitive.isNumber() ? jsonPrimitive.getAsDouble() : jsonPrimitive.getAsString();
        }
        return null;
    }

    private static Object getData(JsonElement values) {
        JsonArray jsonArray = values.getAsJsonArray();
        for (JsonElement jsonElement : jsonArray) {
            for (JsonElement value : jsonElement.getAsJsonArray()) {
                JsonPrimitive jsonPrimitive = value.getAsJsonPrimitive();
                return jsonPrimitive.isNumber() ? jsonPrimitive.getAsDouble() : jsonPrimitive.getAsString();
            }
        }
        return null;
    }

    public static Table parse(WorkbookTable workbookTable, String worksheetName) {
        return (workbookTable != null)
                ? new Table(workbookTable.id, workbookTable.name, worksheetName)
                : null;
    }

    public static JsonElement convertRangeIntoFlatRange(JsonElement values) {
        List<Integer> numbers = convertRangeIntoFlatList(values);
        return Constants.GSON.toJsonTree(numbers);
    }

    public static List<Integer> convertRangeIntoFlatList(JsonElement values) {
        List<Integer> numbers = new ArrayList<>();
        JsonArray jsonArray = values.getAsJsonArray();
        for (JsonElement jsonElement : jsonArray) {
            List<Integer> numberList = Constants.GSON.fromJson(jsonElement, List.class);
            numbers.addAll(numberList);
        }
        return numbers;
    }


    public static FreeForm convertToFreeForm(String formula, JsonElement formulaItem) {
        FreeForm freeForm = new FreeForm();
        if (formulaItem.isJsonPrimitive()) {
            JsonPrimitive jsonPrimitive = formulaItem.getAsJsonPrimitive();
            if (jsonPrimitive.isNumber()) {
                freeForm.put(formula, NUMBER, jsonPrimitive.getAsInt());
            } else {
                freeForm.put(formula, TEXT, jsonPrimitive.getAsString());
            }
        }
        return freeForm;
    }
}
