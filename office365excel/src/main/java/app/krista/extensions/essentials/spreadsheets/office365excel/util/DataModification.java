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

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.model.base.EntityValue;
import app.krista.model.entity.EntityAttributeField;
import app.krista.model.entity.EntityDefinition;
import com.microsoft.graph.models.User;
import org.jetbrains.annotations.NotNull;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;
import java.text.DecimalFormat;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.DateUtil.formatDate;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.DateUtil.formatTime;
import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.FieldConstants.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.UTC;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.MAILBOX_SETTINGS;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.WINDOWS_TO_IANA;

public class DataModification {

    private static final Logger LOGGER = LoggerFactory.getLogger(DataModification.class);
    private static final ThreadLocal<String> USER_TIMEZONE = new ThreadLocal<>();


    /**
     * Converts a list of {@link EntityValue} objects into a list of maps, excluding a specified column if requested.
     *
     * @param list                        The list of {@link EntityValue} objects to convert.
     * @param excludeColumns              The column to exclude from the resulting maps. If null, no column is excluded.
     * @param entityDefinition            The entity definition for field metadata.
     * @param graphServiceClientProvider  The Graph Service Client provider to retrieve user timezone (optional, can be null).
     * @return A list of maps representing the fields of each {@link EntityValue} object in the input list.
     */
    public static List<Map<String, Object>> fromListOfEntities(List<EntityValue> list, List<String> excludeColumns,
                                                                EntityDefinition entityDefinition,
                                                                GraphServiceClientProvider graphServiceClientProvider) {
        if (list == null || list.isEmpty()) {
            throw new IllegalArgumentException("No data to add to desired Sheet");
        }

        // Get user timezone from Microsoft Graph API
        String userTimeZone = getUserTimeZone(graphServiceClientProvider);
        USER_TIMEZONE.set(userTimeZone);

        try {
            List<Map<String, Object>> listOfMap = new ArrayList<>();
            for (EntityValue entityValue : list) {
                Map<String, Object> fields = dataFormat(entityDefinition, entityValue.getFields());
                if (excludeColumns != null && !excludeColumns.isEmpty()) {
                    excludeColumns.forEach(fields.keySet()::remove);
                }
                listOfMap.add(fields);
            }

            // Maintain order based on EntityDefinition, similar to Excel code
            return makeMapsSameSizeWithEntityOrder(listOfMap, entityDefinition, excludeColumns);
        } finally {
            // Clean up thread-local to prevent memory leaks
            USER_TIMEZONE.remove();
        }
    }

    /**
     * Gets user's timezone from Microsoft Graph API.
     *
     * @param provider Graph Service Client provider
     * @return Java IANA timezone ID (e.g., "Asia/Kolkata"), defaults to "UTC" if unable to retrieve
     */
    private static String getUserTimeZone(GraphServiceClientProvider provider) {
        if (provider == null) {
            return UTC;
        }

        try {
            User user = provider.getGraphServiceClient()
                    .me()
                    .buildRequest()
                    .select(MAILBOX_SETTINGS)
                    .get();

            String windowsTimeZone = user.mailboxSettings.timeZone;
            String javaTimeZone = WINDOWS_TO_IANA.getOrDefault(windowsTimeZone, UTC);
            LOGGER.info("User timezone: {} -> {}", windowsTimeZone, javaTimeZone);
            return javaTimeZone;
        } catch (IOException | RuntimeException e) {
            LOGGER.warn("Failed to get user timezone, using UTC: {}", e.getMessage());
            return UTC;
        }
    }

    /**
     * Backward compatibility method - calls the new method with null graphServiceClientProvider.
     * This will use UTC as the default timezone.
     *
     * @deprecated Use {@link #fromListOfEntities(List, List, EntityDefinition, GraphServiceClientProvider)} instead
     */
    @Deprecated
    public static List<Map<String, Object>> fromListOfEntities(List<EntityValue> list, List<String> excludeColumns, EntityDefinition entityDefinition) {
        return fromListOfEntities(list, excludeColumns, entityDefinition, null);
    }

    private static Map<String, Object> dataFormat(EntityDefinition entityDefinition, Map<String, Object> fields) {
        for (Map.Entry<String, Object> entry : fields.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();

            EntityAttributeField entityAttributeField = entityDefinition.getAttributeFields().get(key);
            if (entityAttributeField != null) {
                Object formattedValue = formatFieldByType(entityAttributeField, value);
                fields.put(key, formattedValue);
            }
        }
        return fields;
    }

    private static Object formatFieldByType(EntityAttributeField entityAttributeField, Object value) {
        String fieldType = entityAttributeField.getFieldType();

        return switch (fieldType) {
            case DATE_FIELD_LOGICAL_NAME -> formatDateField(entityAttributeField, value);
            case TIME_FIELD_LOGICAL_NAME -> formatTimeField(value);
            case RICH_TEXT_FIELD_LOGICAL_NAME, TEXT_FIELD_LOGICAL_NAME -> formatTextField(value);
            case DATE_TIME_RANGE_LOGICAL_NAME, DATE_RANGE_LOGICAL_NAME ->
                    formatDateRangeField(entityAttributeField, value);
            case PHONE_FIELD_LOGICAL_NAME -> formatPhoneNumberField(value);
            case CURRENCY_FIELD_LOGICAL_NAME -> formatCurrencyFieldWithStrategy(entityAttributeField, value);
            case NUMBER_FIELD_LOGICAL_NAME -> formatNumberField(entityAttributeField, value);
            case UNIT_FIELD_LOGICAL_NAME -> formatUnitFieldWithStrategy(entityAttributeField, value);
            case PERCENTAGE_FIELD_LOGICAL_NAME -> formatPercentageFieldWithStrategy(value);
            default -> value;
        };
    }

    private static Object formatDateField(EntityAttributeField entityAttributeField, Object value) {
        Boolean includeTime = (Boolean) entityAttributeField.getAttributeValues().get(INCLUDE_TIME_OF_DAY_KEY);
        String dateFormat = (String) entityAttributeField.getAttributeValues().get(DATE_FORMAT_KEY);
        // Use user's timezone from Microsoft Graph API (stored in thread-local)
        String timezone = USER_TIMEZONE.get();
        return formatDate((includeTime != null && includeTime), (Long) value, dateFormat, timezone);
    }

    private static Object formatTimeField(Object value) {
        // Use user's timezone from Microsoft Graph API (stored in thread-local)
        String timezone = USER_TIMEZONE.get();
        return formatTime((Long) value, timezone);
    }

    private static Object formatTextField(Object value) {
        return value != null ? value.toString() : EMPTY_STRING;
    }

    private static Object formatDateRangeField(EntityAttributeField entityAttributeField, Object value) {
        if (!(value instanceof Map)) {
            return EMPTY_STRING;
        }

        Map<String, Object> rangeMap = (Map<String, Object>) value;
        Object fromValue = rangeMap.get(FROM_KEY);
        Object toValue = rangeMap.get(TO_KEY);

        if (fromValue == null || toValue == null) {
            return EMPTY_STRING;
        }

        String dateFormat = (String) entityAttributeField.getAttributeValues().get(DATE_FORMAT_KEY);
        Boolean includeTime = (Boolean) entityAttributeField.getAttributeValues().get(INCLUDE_TIME_OF_DAY_KEY);

        // Use user's timezone from Microsoft Graph API (stored in thread-local)
        String timezone = USER_TIMEZONE.get();
        String fromFormatted = formatDate((includeTime != null && includeTime), (Long) fromValue, dateFormat, timezone);
        String toFormatted = formatDate((includeTime != null && includeTime), (Long) toValue, dateFormat, timezone);

        return fromFormatted + DATE_RANGE_SEPARATOR + toFormatted;
    }

    private static Object formatPhoneNumberField(Object value) {
        if (!(value instanceof Map)) {
            return EMPTY_STRING;
        }

        Map<String, Object> phoneMap = (Map<String, Object>) value;
        Object countryCode = phoneMap.get(COUNTRY_CODE_KEY);
        Object nationalNumber = phoneMap.get(NATIONAL_NUMBER_KEY);

        if (countryCode == null || nationalNumber == null) {
            return EMPTY_STRING;
        }

        return PHONE_NUMBER_PREFIX + countryCode + PHONE_NUMBER_SEPARATOR + nationalNumber;
    }

    private static Object formatCurrencyFieldWithStrategy(EntityAttributeField entityAttributeField, Object value) {
        String currencyFormat = EMPTY_STRING;
        if (!Objects.equals(String.valueOf(value), EMPTY_STRING)) {
            if (entityAttributeField.getAttributeValues() != null) {
                currencyFormat = entityAttributeField.getAttributeValues().get(FORMAT_PATTERN) != null ?
                        entityAttributeField.getAttributeValues().get(FORMAT_PATTERN).toString() : "";
            }

            if (!currencyFormat.isEmpty()) {
                double currencyValue = Double.parseDouble(String.valueOf(value));
                String cleanFormat = currencyFormat.replace("\"", "");
                DecimalFormat decFormat = new DecimalFormat(cleanFormat);
                return decFormat.format(currencyValue);
            }

            double currencyValue = Double.parseDouble(String.valueOf(value));
            return DEFAULT_CURRENCY_SYMBOL + String.format(DEFAULT_CURRENCY_FORMAT, currencyValue);
        } else {
            return EMPTY_STRING;
        }
    }

    private static Object formatNumberField(EntityAttributeField entityAttributeField, Object value) {
        if (value == null) return EMPTY_STRING;

        String formatPattern = entityAttributeField.getAttributeValues().get(FORMAT_PATTERN) != null ?
                entityAttributeField.getAttributeValues().get(FORMAT_PATTERN).toString() : "";

        if (!formatPattern.isEmpty() && value instanceof Number) {
            return String.format(formatPattern.replace("#", "%"), ((Number) value).doubleValue());
        }

        return value.toString();
    }

    private static Object formatUnitFieldWithStrategy(EntityAttributeField entityAttributeField, Object value) {
        if (!Objects.equals(String.valueOf(value), EMPTY_STRING)) {
            String formatPattern = EMPTY_STRING;
            if (entityAttributeField.getAttributeValues() != null) {
                formatPattern = entityAttributeField.getAttributeValues().get(FORMAT_PATTERN) != null ?
                        entityAttributeField.getAttributeValues().get(FORMAT_PATTERN).toString() : "";
            }

            if (!formatPattern.isEmpty()) {
                double unitValue = Double.parseDouble(String.valueOf(value));
                String cleanFormat = formatPattern.replace("\"", "");
                DecimalFormat decFormat = new DecimalFormat(cleanFormat);
                return decFormat.format(unitValue);
            }

            double unitValue = Double.parseDouble(String.valueOf(value));
            return String.valueOf(unitValue);
        } else {
            return EMPTY_STRING;
        }
    }

    private static Object formatPercentageFieldWithStrategy(Object value) {
        if (!Objects.equals(String.valueOf(value), EMPTY_STRING)) {
            double percentageValue = Double.parseDouble(String.valueOf(value));
            DecimalFormat decFormat = new DecimalFormat(PERCENTAGE_VALUE_FORMAT);
            return decFormat.format(percentageValue * 100) + PERCENTAGE_SIGN;
        } else {
            return EMPTY_STRING;
        }
    }

    /**
     * Ensures that a list of maps have the same keys in the same order by adding missing keys with null values.
     *
     * @param inputList The list of maps to standardize.
     * @return A list of maps with uniform keys and order.
     */
    private static List<Map<String, Object>> makeMapsSameSizeWithEntityOrder(List<Map<String, Object>> inputList, EntityDefinition entityDefinition, List<String> excludeColumns) {
        if (inputList == null || inputList.isEmpty()) {
            return Collections.emptyList();
        }

        // Get ordered keys from EntityDefinition (similar to Excel RowUtils approach)
        List<String> orderedKeys = entityDefinition.getAttributeFields().keySet().stream()
                .filter(key -> excludeColumns == null || !excludeColumns.contains(key))
                .toList();

        // Update each map to include all keys in the same order
        inputList.forEach(map -> {
            Map<String, Object> updatedMap = new LinkedHashMap<>();
            for (String key : orderedKeys) {
                updatedMap.put(key, map.getOrDefault(key, ""));
            }
            map.clear();
            map.putAll(updatedMap);
        });

        return inputList;
    }


    /**
     * Converts a list of maps into a list of lists, where each inner list contains string representations of map values.
     *
     * @param listOfMap The list of maps to convert.
     * @return A list of lists representing the map values as strings.
     */
    public static List<List<String>> convertToListOfLists(List<Map<String, Object>> listOfMap) {
        List<String> header = getProperHeader(listOfMap); // Get the proper header
        List<List<String>> collect = new ArrayList<>();
        collect.add(header); // Add the header as the first element
        collect.addAll(listOfMap.stream()
                .map(map -> map.values().stream()
                        .map(Object::toString)
                        .collect(Collectors.toList()))
                .toList());
        return collect;
    }


    /**
     * Retrieves the proper header (keys) from a list of maps, based on the map with the most keys.
     *
     * @param dataList The list of maps to determine the header from.
     * @return A list of strings representing the header (keys) in the correct order.
     */
    public static List<String> getProperHeader(List<Map<String, Object>> dataList) {
        Map<String, Object> mapWithMostKeys = Collections.emptyMap();
        for (Map<String, Object> dataMap : dataList) {
            if (dataMap.size() > mapWithMostKeys.size()) {
                mapWithMostKeys = dataMap;
            }
        }
        return new ArrayList<>(mapWithMostKeys.keySet());
    }


    /**
     * Extracts a sublist of lists from a given list of lists, starting from the specified index.
     *
     * @param listOfLists The original list of lists.
     * @param startIndex  The starting index of the sublist.
     * @return A sublist of lists starting from the specified index.
     * @throws IndexOutOfBoundsException If the startIndex is negative or greater than or equal to the size of the input list.
     */
    public static List<List<String>> convertToListOfListsFromIndex(List<List<String>> listOfLists, int startIndex) {
        if (startIndex < 0 || startIndex >= listOfLists.size()) {
            return Collections.emptyList();
        }
        return listOfLists.subList(startIndex, listOfLists.size());
    }

    @NotNull
    public static Map<String, Object> getResultMap(String pdfFileId, app.krista.model.base.File pdfFile) {
        Map<String, Object> pdfFileAndFileIdMap = new HashMap<>();
        pdfFileAndFileIdMap.put("PDF File", pdfFile);
        pdfFileAndFileIdMap.put("File Id", pdfFileId);
        return pdfFileAndFileIdMap;
    }

}