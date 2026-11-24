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

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.FileService;
import app.krista.model.base.EntityValue;
import app.krista.model.base.File;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * This class is util class and is designed to help others for validating String.
 */
@Service
public final class StringUtil {

    private final KristaMediaClient kristaMediaClient;

    @Inject
    public StringUtil(KristaMediaClient kristaMediaClient) {
        this.kristaMediaClient = kristaMediaClient;
    }

    /**
     * @param string to validate
     * @return true, if string is null or empty or having whitespaces only else return false.
     */
    public static boolean isNullOrEmpty(String string) {
        return string == null || string.isBlank();
    }

    public static String emptyIfNull(String string) {
        return string == null ? "" : string;
    }

    public static boolean isExcelFile(String driveItem) {
        return driveItem.endsWith("xlsx") || driveItem.endsWith("xls") || driveItem.endsWith("csv") || driveItem.endsWith("tsv");
    }


    public List<String> getStringList(List<String> rows, List<File> rowsFromCsv)
            throws IOException {

        List<String> rowsToAdd = rows;

        if (rows == null || rows.isEmpty()) {
            FileService fileService = new FileService(kristaMediaClient);
            rowsToAdd = fileService.getRows(rowsFromCsv);
        }

        return rowsToAdd;
    }

    public static List<String> appendCommas(List<String> inputList) {
        int maxCommas = 0;

        // Find the string with the maximum number of commas
        for (String str : inputList) {
            int commaCount = countCommas(str);
            if (commaCount > maxCommas) {
                maxCommas = commaCount;
            }
        }

        // Append commas to other strings to match the maximum number
        for (int i = 0; i < inputList.size(); i++) {

            String currentString = inputList.get(i);
            int commaCount = countCommasIgnoringQuotes(currentString);
            int commasToAdd = maxCommas - commaCount;
            String appendedString = currentString + repeatCharacter(commasToAdd);
            inputList.set(i, appendedString);
        }

        return inputList;
    }

    private static int countCommas(String str) {
        int count = 0;
        for (char c : str.toCharArray()) {
            if (c == ',') {
                count++;
            }
        }
        return count;
    }

    /**
     * Counts the number of commas in a string, ignoring commas that appear within quotes.
     *
     * @param str The input string to be processed.
     * @return The count of commas present in the string, excluding commas within quotes.
     */
    private static int countCommasIgnoringQuotes(String str) {
        int count = 0;
        boolean insideQuotes = false;

        for (char c : str.toCharArray()) {
            if (c == '\"') {
                insideQuotes = !insideQuotes; // Toggle the insideQuotes flag when encountering a quote
            } else if (c == ',' && !insideQuotes) {
                count++;
            }
        }

        return count;
    }

    private static String repeatCharacter(int times) {

        return String.valueOf(',').repeat(Math.max(0, times));
    }

    public static int findMaxListSize(List<String> inputList) {
        int maxSize = 0;

        for (String row : inputList) {
            List<String> splitList = Arrays.stream(row.split(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"
                            , -1))
                    .map(e -> e.replaceAll("\"", ""))
                    .toList();

            int currentSize = splitList.size();
            if (currentSize > maxSize) {
                maxSize = currentSize;
            }
        }

        return maxSize;
    }

    public static boolean checkColonPresence(String input) {
        return input.contains(":");
    }

    public static String extractAlphabetAfterColon(String input) {
        int colonIndex = input.indexOf(":");
        if (colonIndex != -1) {
            String alphabetSubstring = input.substring(colonIndex + 1).trim();
            StringBuilder alphabetBuilder = new StringBuilder();

            for (int i = 0; i < alphabetSubstring.length(); i++) {
                char ch = alphabetSubstring.charAt(i);
                if (Character.isLetter(ch)) {
                    alphabetBuilder.append(ch);
                } else {
                    break;
                }
            }

            return alphabetBuilder.toString();
        }

        return "";
    }

    public static void removeEmptyStringsFromEnd(List<String> stringList) {
        int lastIndex = stringList.size() - 1;

        // Iterate over the list from the end
        for (int i = lastIndex; i >= 0; i--) {
            // Get the current string
            String currentString = stringList.get(i);

            // Check if it is an empty string
            if (currentString.isEmpty()) {
                // Remove the empty string
                stringList.remove(i);
            } else {
                // Exit the loop as we encountered a non-empty string
                break;
            }
        }
    }

    public String calculateAddressForInsert(Double rowIndex, List<String> rows,
                                            List<File> rowsFromCsv) {
        List<String> rowsToAdd = rows;

        try {
            // If the initial rows list is empty or null, retrieve rows from CSV files using FileService
            if (rows == null || rows.isEmpty()) {
                FileService fileService = new FileService(kristaMediaClient);
                rowsToAdd = fileService.getRows(rowsFromCsv);
            }
            String firstColumn = "A";
            String lastColumn = "XFD";
            int finalRow = rowIndex.intValue() + rowsToAdd.size() - 1;
            // Return the address range in the format "A(rowIndex):XFD(finalRow)"
            return firstColumn + rowIndex.intValue() + ":" + lastColumn + finalRow;
        } catch (IOException exception) {
            return null;
        }

    }

    public static String calculateAddressForInsert(Double rowIndex, List<String> rows) {

        int maxStringSize = findMaxListSize(rows);
        String firstColumn = "A";

        int lastCellNumber = rowIndex.intValue() + rows.size() - 1;
        String lastCellID = getColumnLabel(maxStringSize) + lastCellNumber;

        return firstColumn + rowIndex.intValue() + ":" + lastCellID;
    }


    /**
     * Converts a column number to its corresponding column label.
     *
     * @param columnNumber The column number to be converted.
     * @return The column label corresponding to the given column number.
     */

    public static String getColumnLabel(int columnNumber) {
        StringBuilder label = new StringBuilder();

        while (columnNumber > 0) {
            int remainder = (columnNumber - 1) % 26;
            char character = (char) (remainder + 'A');
            label.insert(0, character);
            columnNumber = (columnNumber - 1) / 26;
        }

        return label.toString();
    }

    public static boolean isDataNullOrEmpty(List<EntityValue> data) {
        return data == null || data.isEmpty();
    }

}
