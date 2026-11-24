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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelSheet;
import app.krista.extensions.essentials.spreadsheets.office365excel.data.formatter.DataFormatter;
import app.krista.extensions.essentials.spreadsheets.office365excel.data.formatter.DataFormatterFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.base.File;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonNull;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.models.WorkbookWorksheet;
import org.jetbrains.annotations.NotNull;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.ParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

public final class Worksheets {
    private static final Logger LOGGER = LoggerFactory.getLogger(Worksheets.class);

    private Worksheets() {
    }

    public static List<ExcelSheet> getWorksheets(List<WorkbookWorksheet> worksheets) {
        List<ExcelSheet> workSheetsList = new ArrayList<>();
        for (WorkbookWorksheet workbookWorksheet : worksheets) {
            ExcelSheet excelSheet = getWorksheet(workbookWorksheet);
            workSheetsList.add(excelSheet);
        }
        return workSheetsList;
    }

    public static ExcelSheet getWorksheet(WorkbookWorksheet worksheet) {
        return new ExcelSheet(worksheet.id, worksheet.name);
    }

    public static WorkbookRange updateWorkbookRange(String worksheetName, WorkbookRange workbookRange, List<String> rows, String customAddress) {
        LOGGER.info("Starting updateWorkbookRange for worksheet: {}, rows count: {}, customAddress: {}",
                worksheetName, rows != null ? rows.size() : 0, customAddress);

        JsonElement jsonValues = new JsonArray();
        JsonElement jsonFormulas = new JsonArray();
        WorkbookRange newWorkbookRange = new WorkbookRange();
        boolean doesNotContainColon = false;
        String address = workbookRange.address;

        // If the address does not contain ":" then add
        if (address != null && !address.contains(Constants.COLON)) {
            address += Constants.COLON + address.split("!")[1];
            doesNotContainColon = true;
        }

        String[] cellRange = address.split(Constants.COLON);
        Integer lastRowNumber = getRowNumber(cellRange[1]);
        int firstRowNumber = lastRowNumber + 1;
        int columnNumber = 1;

        // --- NEW: force column count based on customAddress ---
        int expectedColumns = columnNumber; // default: just input width

        if (customAddress != null && customAddress.contains(":")) {
            try {
                String[] customCellRange = customAddress.split(":");
                String startCell = customCellRange[0];
                String endCell = customCellRange[1];

                int startCol = getColumnNumber(startCell);
                int endCol = getColumnNumber(endCell);

                expectedColumns = endCol - startCol + 1;
                LOGGER.info("Forced expected columns based on customAddress: {}", expectedColumns);
            } catch (Exception e) {
                LOGGER.warn("Could not compute expectedColumns from customAddress: {}", customAddress, e);
                expectedColumns = columnNumber; // fallback
            }
        }
        // --- END NEW ---

        boolean hasFormulas = false;
        for (int rowIndex = 0; rowIndex < Objects.requireNonNull(rows).size(); rowIndex++) {
            String row = rows.get(rowIndex);
            LOGGER.debug("Processing row {}: {}", rowIndex, row);

            // Split the row into cells, preserving quoted commas
            List<String> rowValues = row.trim().startsWith("=")
                    ? List.of(row.trim())
                    : Arrays.stream(row.split(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", -1)).toList();

            JsonArray valueArray = new JsonArray();
            JsonArray formulaArray = new JsonArray();

            for (int cellIndex = 0; cellIndex < rowValues.size(); cellIndex++) {
                String processedCell = rowValues.get(cellIndex).trim();

                if (processedCell.startsWith("\"") && processedCell.endsWith("\"")) {
                    processedCell = processedCell.substring(1, processedCell.length() - 1);
                    LOGGER.trace("Cell [{},{}] - removed quotes: {}", rowIndex, cellIndex, processedCell);
                }

                if (processedCell.startsWith("=HYPERLINK")) {
                    LOGGER.debug("Found HYPERLINK formula: {}", processedCell);
                    processedCell = processedCell.replace("\"\"", "\"");
                    LOGGER.debug("Processed HYPERLINK formula: {}", processedCell);
                    hasFormulas = true;
                    formulaArray.add(processedCell);
                    valueArray.add("");
                } else if (processedCell.startsWith("=")) {
                    hasFormulas = true;
                    formulaArray.add(processedCell);
                    valueArray.add("");
                } else {
                    formulaArray.add(JsonNull.INSTANCE);
                    valueArray.add(processedCell);
                }
            }

            // --- NEW: force column number to match input or customAddress ---
            columnNumber = Math.max(expectedColumns, rowValues.size());

            LOGGER.debug("Updated column number to: {}", columnNumber);

            // Pad if needed
            while (valueArray.size() < columnNumber) {
                valueArray.add("");
                formulaArray.add(JsonNull.INSTANCE);
            }

            jsonValues.getAsJsonArray().add(valueArray);
            jsonFormulas.getAsJsonArray().add(formulaArray);

            lastRowNumber += 1;
            LOGGER.debug("Updated last row number to: {}", lastRowNumber);
        }

        // This is needed if the sheet is blank or newly created
        if (doesNotContainColon) {
            firstRowNumber -= 1;
            lastRowNumber -= 1;
        }

        WorkbookRange result = getWorkbookRange(
                worksheetName, customAddress, firstRowNumber, columnNumber, lastRowNumber, newWorkbookRange, jsonValues);

        if (hasFormulas) {
            result.formulas = jsonFormulas;
            result.values = null; // force Graph to take formulas only
        } else {
            result.values = jsonValues;
        }

        LOGGER.info("Address final Before Patch {}", result.address);
        LOGGER.info("Values{}", jsonValues);
        LOGGER.info("Formulas{}", jsonFormulas);

        return result;
    }

    public static int getColumnNumber(String cell) {
        StringBuilder colRef = new StringBuilder();
        for (char c : cell.toCharArray()) {
            if (Character.isLetter(c)) {
                colRef.append(c);
            } else {
                break; // stop at first digit
            }
        }
        String colLetters = colRef.toString().toUpperCase();

        int colNumber = 0;
        for (int i = 0; i < colLetters.length(); i++) {
            colNumber *= 26;
            colNumber += colLetters.charAt(i) - 'A' + 1;
        }
        return colNumber;
    }

    public static Set<String> find(Integer rowNumber, JsonArray row, String lookupValue) {
        char initialColumn = Constants.A;
        Set<String> cellIDs = new TreeSet<>();
        char exactColumn;
        for (int columnNumber = 0; columnNumber < row.size(); columnNumber++) {
            if (lookupValue.equalsIgnoreCase(row.get(columnNumber).getAsString())) {
                exactColumn = (char) (initialColumn + columnNumber); // A+1 results B
                String cellID = String.valueOf(exactColumn) + rowNumber;
                cellIDs.add(cellID);
            }
        }
        return cellIDs;
    }

    public static Integer getRowNumber(String cellID) {
        // \d Matches any digit character (0-9)
        // {1,} Match 1 or more of the preceding token
        // $ Matches the end of the string
        Pattern p = Pattern.compile("\\d{1,}$");
        Matcher m = p.matcher(cellID.trim());
        if (m.find()) {
            return Integer.parseInt(m.group());
        }
        return null;
    }

    public static boolean isValidCellId(String cellId) {
        Pattern p = Pattern.compile("^[a-zA-Z]{1,3}\\d{1,}$");
        Matcher m = p.matcher(cellId.trim());

        String group = null;
        if (m.find()) {
            group = m.group();
        }

        return group != null && !group.isBlank();
    }

    public static boolean isValidColumnRange(String columnRange) {
        String[] split = columnRange.trim().split(Constants.COLON);
        return 2 == split.length && isValidColumnName(split[0]) && isValidColumnName(split[1]);
    }

    private static boolean isValidColumnName(String columnName) {
        return columnName != null && columnName.matches("^[a-zA-Z]*$");
    }

    public static boolean isValidRange(String cellRange) {
        String[] split = cellRange.trim().split(Constants.COLON);
        return 2 == split.length && isValidCellId(split[0]) && isValidCellId(split[1]);
    }

    public static double getSum(JsonElement jsonElement) {
        double total = 0;
        for (JsonElement object : jsonElement.getAsJsonArray()) {
            try {
                total += Double.parseDouble(object.getAsString());
            } catch (NumberFormatException ignore) {
                // Ignoring in such cases as some entries might not be parsable to double
            }
        }
        return total;
    }

    public static boolean isRowDataAvailable(List<String> rows, List<File> rowsFromCsvFile) {
        boolean rowsAvailable = rows != null && !rows.isEmpty();
        boolean csvFilesAvailable = rowsFromCsvFile != null && !rowsFromCsvFile.isEmpty();

        return rowsAvailable || csvFilesAvailable;
    }

    public static List<Object> createKeyColumn(JsonArray worksheetData, JsonArray valueTypes, JsonArray dataFormats, int keyColumnIndex) {
        List<Object> keyColumnList = new ArrayList<>();
        for (int index = 0; index < worksheetData.size(); index++) {
            JsonArray rowData = worksheetData.get(index).getAsJsonArray();
            JsonArray rowValuesType = valueTypes.get(index).getAsJsonArray();
            JsonArray rowDataFormats = dataFormats.get(index).getAsJsonArray();
            String cellValue = rowData.get(keyColumnIndex).getAsString();
            DataFormatter dataFormatter = DataFormatterFactory.create(rowValuesType.get(keyColumnIndex).getAsString(), rowDataFormats.get(keyColumnIndex).getAsString());
            Object value;
            try {
                value = dataFormatter.format(cellValue);
            } catch (ParseException cause) {
                throw new RuntimeException("Exception occurred while parsing " + cellValue + ". " + cause.getMessage());
            }
            keyColumnList.add(value);
        }
        return keyColumnList;
    }

    private static String intToLetter(int Int) {
        if (Int < 27) {
            return Character.toString((char) (Int + 64));
        } else {
            if (Int % 26 == 0) {
                return intToLetter((Int / 26) - 1) + intToLetter(((Int - 1) % 26 + 1));
            } else {
                return intToLetter(Int / 26) + intToLetter(Int % 26);
            }
        }
    }

    public static String[] getSequencedAlphabetsForColumns(int columnCount) {

        String[] sequencedAlphabets = new String[columnCount];

        for (int i = 0; i < columnCount; i++) {
            sequencedAlphabets[i] = intToLetter(i + 1);
        }
        return sequencedAlphabets;
    }

    public static List getWorksheetDataFromRange(JsonArray worksheetData, int startRangeColumn, int endRangeColumn) {

        List<JsonArray> worksheetRangeData = new ArrayList<>();
        for (JsonElement worksheetRow : worksheetData) {
            JsonArray rowData = getWorkSheetRowForRange(startRangeColumn, endRangeColumn, (JsonArray) worksheetRow);
            worksheetRangeData.add(rowData);
        }
        return worksheetRangeData;
    }

    @NotNull
    private static JsonArray getWorkSheetRowForRange(int startRangeColumn, int endRangeColumn, JsonArray worksheetRow) {
        JsonArray rowData = new JsonArray();
        for (int iterator = startRangeColumn; iterator <= endRangeColumn; iterator++) {
            String str = worksheetRow.get(iterator).getAsString();
            rowData.add(str);
        }
        return rowData;
    }

    public static Map<String, Double> findMinMaxValueInWorksheetRange(List<JsonArray> rangeWorksheetData) {
        DecimalFormat df = new DecimalFormat(
                "#,#.######",
                DecimalFormatSymbols.getInstance(Locale.ENGLISH));

        List<Double> values = rangeWorksheetData.stream().map(element -> element.getAsJsonArray().getAsString()).map(element -> {
            try {
                return df.parse(element).doubleValue();
            } catch (ParseException cause) {
                return null;
            }
        }).toList();

        Optional<Double> optionalMax = values.stream().filter(Objects::nonNull).max(Double::compare);
        double maxElement = optionalMax.orElse(null);
        Optional<Double> optionalMin = values.stream().filter(Objects::nonNull).min(Double::compare);
        double minElement = optionalMin.orElse(null);
        OptionalInt maxValueRowNumber = IntStream.range(0, values.size())
                .filter(i -> values.get(i).equals(maxElement)).findFirst();
        OptionalInt minValueRowNumber = IntStream.range(0, values.size())
                .filter(i -> values.get(i).equals(minElement)).findFirst();
        return getMinMaxValueInformation(maxValueRowNumber.isPresent() ? maxValueRowNumber.getAsInt() : 0,
                minValueRowNumber.isPresent() ? minValueRowNumber.getAsInt() : 0,
                maxElement, minElement);
    }

    @NotNull
    private static Map<String, Double> getMinMaxValueInformation(int maxValueRowNumber, int minValueRowNumber, double maxElement, double minElement) {
        Map<String, Double> minMaxValueInformation = new HashMap<>();
        minMaxValueInformation.put(Constants.MAX_VALUE, maxElement);
        minMaxValueInformation.put(Constants.MAX_VALUE_ROW_INDEX, (double) maxValueRowNumber);

        minMaxValueInformation.put(Constants.MIN_VALUE, minElement);
        minMaxValueInformation.put(Constants.MIN_VALUE_ROW_INDEX, (double) minValueRowNumber);

        return minMaxValueInformation;
    }


    public static WorkbookRange updateWorkbookRangeForEntity(String worksheetName, WorkbookRange workbookRange,
                                                             List<List<String>> rows, String customAddress) {
        JsonElement jsonValues = new JsonArray();
        WorkbookRange newWorkbookRange = new WorkbookRange();
        boolean doesNotContainColon = false;
        String address = workbookRange.address;
        //if the address does not contain ":" then add
        if (address != null && !address.contains(Constants.COLON)) {
            address += Constants.COLON + address.split("!")[1];
            doesNotContainColon = true;
        }
        String[] cellRange = address.split(Constants.COLON);
        Integer lastRowNumber = getRowNumber(cellRange[1]);
        int firstRowNumber = lastRowNumber + 1;
        int columnNumber = 1;
        for (List<String> row : rows) {
            columnNumber = Math.max(columnNumber, row.size());
            JsonElement dataArray = Constants.GSON.toJsonTree(row);
            JsonArray jsonArray = jsonValues.getAsJsonArray();
            jsonArray.add(dataArray);
            lastRowNumber += 1;
        }
        // This is needed if the sheet is blank or newly created
        // if blank then A1 should be starting cell and Not A2
        if (doesNotContainColon) {
            firstRowNumber -= 1;
            lastRowNumber -= 1;
        }

        return getWorkbookRange(worksheetName, customAddress, firstRowNumber, columnNumber,
                lastRowNumber, newWorkbookRange, jsonValues);

    }

    /**
     * Constructs a WorkbookRange object based on provided parameters and JSON values.
     * The resulting range can be used to manipulate a worksheet within a workbook.
     *
     * @param worksheetName    The name of the target worksheet.
     * @param customAddress    Custom address (if provided) for the range.
     * @param firstRowNumber   The row number of the first cell in the range.
     * @param columnNumber     The column number of the first cell in the range.
     * @param lastRowNumber    The row number of the last cell in the range.
     * @param newWorkbookRange The WorkbookRange object to be configured.
     * @param jsonValues       JSON values to be associated with the range.
     * @return The configured WorkbookRange object.
     */

    private static WorkbookRange getWorkbookRange(String worksheetName, String customAddress,
                                                  Integer firstRowNumber, Integer columnNumber, Integer lastRowNumber,
                                                  WorkbookRange newWorkbookRange, JsonElement jsonValues) {
        String firstCellID = "A" + firstRowNumber;
        LOGGER.info("Column Number {}", columnNumber);
        String lastCellID = StringUtil.getColumnLabel(columnNumber) + lastRowNumber;
        newWorkbookRange.address = customAddress == null ? worksheetName + "!" + firstCellID + ":" + lastCellID : worksheetName + "!" + customAddress;
        LOGGER.info("Address final Before Patch {}", newWorkbookRange.address);
        newWorkbookRange.values = jsonValues;
        LOGGER.info("Values{}", Constants.GSON.toJson(jsonValues));
        return newWorkbookRange;
    }

    public static List<String> search(WorkbookRange wbRange, String lookupValue) {
        if (wbRange == null || wbRange.text == null) {
            return Collections.emptyList();
        }
        JsonArray rows = wbRange.text.getAsJsonArray();
        Set<String> cellIds = new TreeSet<>();
        for (int rowNumber = 0; rowNumber < rows.size(); rowNumber++) {
            JsonArray rowData = rows.get(rowNumber).getAsJsonArray();
            cellIds.addAll(Worksheets.find(rowNumber + 1, rowData, lookupValue));
        }
        return new ArrayList<>(cellIds);
    }

}
