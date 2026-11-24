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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.attributes.Print;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.model.field.NamedField;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes.*;
import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.FieldConstants.*;

/**
 * Utility class providing static factory methods for creating {@link NamedField} instances
 * with predefined field types for use in Office 365 calendar extension responses.
 * <p>
 * This class simplifies the creation of commonly used field types such as text fields, switch fields,
 * number fields, boolean fields, date fields, rich text fields, file fields, and pick-one fields,
 * ensuring consistency and reducing boilerplate code throughout the extension.
 * </p>
 */
public class NamedFieldFactory {

    /**
     * Creates a switch field.
     */
    public static NamedField createSwitchField(String fieldName) {
        return new NamedField(fieldName, SWITCH_FIELD_LOGICAL_NAME, Collections.emptyMap(), Collections.emptyMap());
    }

    /**
     * Creates a field with specified type and attributes.
     */
    public static NamedField createFieldWithAttributes(String fieldName, String fieldType, Map<String, Object> attributes) {
        // Handle catalog request list format first (e.g., "[ Email ]", "[ Text ]", "[ File ]")
        if (fieldType.startsWith(LIST_BRACKET_START) && fieldType.endsWith(LIST_BRACKET_END)) {
            String innerType = fieldType.substring(1, fieldType.length() - 1).trim();
            return createListFieldWithAttributes(fieldName, innerType, attributes);
        }

        // Handle list types with inner type specification (e.g., "List<Email>", "List[Date]", etc.)
        if (fieldType.startsWith(LIST_ANGLE_START) && fieldType.endsWith(LIST_BRACKET_END)) {
            String innerType = fieldType.substring(5, fieldType.length() - 1);
            return createListFieldWithAttributes(fieldName, innerType, attributes);
        }

        if (fieldType.startsWith(LIST_SQUARE_START) && fieldType.endsWith(LIST_BRACKET_END)) {
            String innerType = fieldType.substring(5, fieldType.length() - 1);
            return createListFieldWithAttributes(fieldName, innerType, attributes);
        }

        // Handle simple list case
        if (fieldType.equals(LIST_FIELD_LOGICAL_NAME) || fieldType.equalsIgnoreCase(LIST_TYPE)) {
            return createListFieldWithAttributes(fieldName, DEFAULT_LIST_INNER_TYPE, attributes);
        }

        // Handle password field special case
        if (fieldType.equals(TEXT_FIELD_LOGICAL_NAME) || fieldType.equalsIgnoreCase(DEFAULT_LIST_INNER_TYPE)) {
            if (fieldName.equalsIgnoreCase(PASSWORD_FIELD_NAME)) {
                Map<String, Object> mergedAttributes = new HashMap<>(attributes);
                mergedAttributes.put(PASSWORD_ATTRIBUTE, true);
                return new NamedField(fieldName, TEXT_FIELD_LOGICAL_NAME, mergedAttributes, Collections.emptyMap());
            }
        }

        String innerFieldLogicalName = "";
        if (fieldType.equalsIgnoreCase(FieldTypes.FILE_FIELD_LOGICAL_NAME)) {
            innerFieldLogicalName = FILE_TYPE_LOWERCASE;
        }
        return createFieldByTypeWithAttributes(fieldName, innerFieldLogicalName, attributes);
    }

    /**
     * Creates a field based on the specified field type with custom attributes.
     */
    public static NamedField createFieldByTypeWithAttributes(String fieldName, String fieldType, Map<String, Object> attributes) {
        return switch (fieldType.toLowerCase()) {
            case DATE_TYPE -> createDateFieldWithAttributes(fieldName, attributes);
            case EMAIL_TYPE -> createEmailFieldWithAttributes(fieldName, attributes);
            case NUMBER_TYPE -> createNumberFieldWithAttributes(fieldName, attributes);
            case BOOLEAN_TYPE -> createBooleanFieldWithAttributes(fieldName, attributes);
            case FILE_TYPE -> createFileFieldWithAttributes(fieldName, attributes);
            case SWITCH_TYPE -> createSwitchFieldWithAttributes(fieldName, attributes);
            case RICHTEXT_TYPE -> createRichTextFieldWithAttributes(fieldName, attributes);
            case PICKONE_TYPE -> createPickOneFieldWithAttributes(fieldName, attributes);
            case PASSWORD_ATTRIBUTE -> {
                // Merge password attributes with provided attributes
                Map<String, Object> mergedAttributes = new HashMap<>(attributes);
                mergedAttributes.put(PASSWORD_ATTRIBUTE, true);
                yield new NamedField(fieldName, TEXT_FIELD_LOGICAL_NAME, mergedAttributes, Map.of());
            }
            case "list" -> createListFieldWithAttributes(fieldName, DEFAULT_LIST_INNER_TYPE, attributes);
            default -> createTextFieldWithAttributes(fieldName, attributes);
        };
    }

    /**
     * Creates a text field with attributes.
     */
    public static NamedField createTextFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        // Get default attributes for text fields and merge with provided attributes
        Map<String, Object> finalAttributes = getDefaultAttributesForFieldType(DEFAULT_LIST_INNER_TYPE);
        if (attributes != null && !attributes.isEmpty()) {
            finalAttributes.putAll(attributes); // Override defaults with provided attributes
        }

        return new NamedField(fieldName, TEXT_FIELD_LOGICAL_NAME, finalAttributes, Collections.emptyMap());
    }

    /**
     * Creates a switch field with attributes.
     */
    public static NamedField createSwitchFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        return new NamedField(fieldName, FieldTypes.SWITCH_FIELD_LOGICAL_NAME, attributes, Collections.emptyMap());
    }

    /**
     * Creates a number field with attributes.
     */
    public static NamedField createNumberFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        return new NamedField(fieldName, FieldTypes.NUMBER_FIELD_LOGICAL_NAME, attributes, Collections.emptyMap());
    }

    /**
     * Creates a boolean field with attributes.
     */
    public static NamedField createBooleanFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        return new NamedField(fieldName, FieldTypes.BOOLEAN_FIELD_LOGICAL_NAME, attributes, Collections.emptyMap());
    }

    /**
     * Creates a date field with attributes.
     */
    public static NamedField createDateFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        return new NamedField(fieldName, FieldTypes.DATE_FIELD_LOGICAL_NAME, attributes, Collections.emptyMap());
    }

    /**
     * Creates a rich text field with attributes.
     */
    public static NamedField createRichTextFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        return new NamedField(fieldName, FieldTypes.RICH_TEXT_FIELD_LOGICAL_NAME, attributes, Collections.emptyMap());
    }

    /**
     * Creates a file field with attributes.
     */
    public static NamedField createFileFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        Map<String, Object> finalAttributes = getDefaultAttributesForFieldType(FILE_TYPE_LOWERCASE);
        Print.prettyPrint(finalAttributes);
        return new NamedField(fieldName, FieldTypes.FILE_FIELD_LOGICAL_NAME, finalAttributes, Collections.emptyMap());
    }

    /**
     * Creates a pick-one field with attributes.
     */
    public static NamedField createPickOneFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        return new NamedField(fieldName, FieldTypes.PICK_ONE_LOGICAL_NAME, attributes, Collections.emptyMap());
    }

    /**
     * Creates an email field with attributes.
     */
    public static NamedField createEmailFieldWithAttributes(String fieldName, Map<String, Object> attributes) {
        return new NamedField(fieldName, FieldTypes.EMAIL_FIELD_LOGICAL_NAME, attributes, Collections.emptyMap());
    }

    /**
     * Creates a list field with custom attributes.
     */
    public static NamedField createListFieldWithAttributes(String fieldName, String innerFieldType, Map<String, Object> attributes) {
        // Map the inner field type to the appropriate logical name
        String innerFieldLogicalName = getLogicalNameForFieldType(innerFieldType);

        // Get default attributes for the inner field type and merge with provided attributes
        Map<String, Object> finalAttributes = getDefaultAttributesForFieldType(innerFieldType);

        // Create the inner field configuration
        Map<String, Object> fieldMap = new HashMap<>();
        fieldMap.put(FIELD_TYPE_KEY, FIELD_TYPE_VALUE);
        fieldMap.put(FIELD_TYPE_ATTR_KEY, innerFieldLogicalName);
        fieldMap.put(ATTRIBUTES_KEY, finalAttributes);

        // This defines that the list holds fields of the specified type
        String fieldMapJson = convertMapToJson(fieldMap);

        // Create the list field attributes
        Map<String, Object> listAttributes = new HashMap<>();
        listAttributes.put(FIELD_KEY, fieldMapJson);
        Print.prettyPrint(listAttributes);

        return new NamedField(fieldName, FieldTypes.LIST_FIELD_LOGICAL_NAME, listAttributes, Collections.emptyMap());
    }


    // ========== Private Helper Methods ==========

    /**
     * Maps a field type string to its corresponding logical name.
     */
    private static String getLogicalNameForFieldType(String fieldType) {
        return switch (fieldType.toLowerCase()) {
            case TEXT_TYPE -> TEXT_FIELD_LOGICAL_NAME;
            case DATE_TYPE -> FieldTypes.DATE_FIELD_LOGICAL_NAME;
            case EMAIL_TYPE -> FieldTypes.EMAIL_FIELD_LOGICAL_NAME;
            case NUMBER_TYPE -> FieldTypes.NUMBER_FIELD_LOGICAL_NAME;
            case BOOLEAN_TYPE -> FieldTypes.BOOLEAN_FIELD_LOGICAL_NAME;
            case FILE_TYPE -> FieldTypes.FILE_FIELD_LOGICAL_NAME;
            case SWITCH_TYPE -> FieldTypes.SWITCH_FIELD_LOGICAL_NAME;
            case RICHTEXT_TYPE -> FieldTypes.RICH_TEXT_FIELD_LOGICAL_NAME;
            case PICKONE_TYPE -> FieldTypes.PICK_ONE_LOGICAL_NAME;
            case PICKMANY_TYPE -> FieldTypes.PICK_MANY_LOGICAL_NAME;
            case PHONE_TYPE -> FieldTypes.PHONE_FIELD_LOGICAL_NAME;
            case ADDRESS_TYPE -> FieldTypes.ADDRESS_FIELD_LOGICAL_NAME;
            case CURRENCY_TYPE -> FieldTypes.CURRENCY_FIELD_LOGICAL_NAME;
            case PERCENTAGE_TYPE -> FieldTypes.PERCENTAGE_FIELD_LOGICAL_NAME;
            case TIME_TYPE -> FieldTypes.TIME_FIELD_LOGICAL_NAME;
            case DATERANGE_TYPE -> FieldTypes.DATE_RANGE_LOGICAL_NAME;
            case DATETIMERANGE_TYPE -> FieldTypes.DATE_TIME_RANGE_LOGICAL_NAME;
            case LOCATION_TYPE -> FieldTypes.LOCATION_FIELD_LOGICAL_NAME;
            default -> TEXT_FIELD_LOGICAL_NAME; // Default fallback
        };
    }

    // Default attribute constants
    private static final Map<String, Object> TEXT_DEFAULT_ATTRIBUTES = Map.of(
        VISUAL_WIDTH_KEY, "S",
        TRIM_KEY, true
    );

    private static final Map<String, Object> DATE_DEFAULT_ATTRIBUTES = Map.of(
        VISUAL_WIDTH_KEY, "S",
        DATE_FORMAT_KEY, "MM/dd/yyyy",
        SHOW_CALENDAR_KEY, true,
        ALLOW_PAST_DATES_KEY, true,
        ALLOW_FUTURE_DATES_KEY, true
    );

    private static final Map<String, Object> FILE_DEFAULT_ATTRIBUTES = Map.of(
        VISUAL_WIDTH_KEY, "S",
        MULTIPLE_FILE_UPLOAD_KEY, true,
        MIN_FILE_SIZE_UNIT_KEY, "KB",
        MAX_FILE_SIZE_UNIT_KEY, "MB",
        FILE_TYPES_KEY, "[\"Document\",\"Image gallery\",\"Camera\",\"Contact\",\"Audio\"]"
    );

    private static final Map<String, Object> DEFAULT_ATTRIBUTES = Map.of(VISUAL_WIDTH_KEY, "S");

    /**
     * Gets default attributes for the specified field type.
     */
    public static Map<String, Object> getDefaultAttributesForFieldType(String fieldType) {
        return switch (fieldType.toLowerCase()) {
            case TEXT_TYPE -> TEXT_DEFAULT_ATTRIBUTES;
            case DATE_TYPE -> DATE_DEFAULT_ATTRIBUTES;
            case FILE_TYPE -> FILE_DEFAULT_ATTRIBUTES;
            default -> DEFAULT_ATTRIBUTES;
        };
    }

    /**
     * Converts a Map to JSON string.
     */
    private static String convertMapToJson(Map<String, Object> map) {
        return Constants.GSON.toJson(map);
    }
}