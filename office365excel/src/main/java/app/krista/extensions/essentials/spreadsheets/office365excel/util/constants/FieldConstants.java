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

package app.krista.extensions.essentials.spreadsheets.office365excel.util.constants;

/**
 * Constants for field creation and configuration in the Office365Excel extension.
 * This class centralizes all field-related string literals and default values.
 */
public final class FieldConstants {

    private FieldConstants() {
    }

    // List parsing constants
    public static final String LIST_BRACKET_START = "[";
    public static final String LIST_BRACKET_END = "]";
    public static final String LIST_ANGLE_START = "List<";
    public static final String LIST_SQUARE_START = "List[";
    public static final String LIST_TYPE = "List";
    public static final String PASSWORD_FIELD_NAME = "Password";
    public static final String PASSWORD_ATTRIBUTE = "password";
    public static final String DEFAULT_LIST_INNER_TYPE = "Text";
    public static final String FILE_TYPE_LOWERCASE = "file";

    // Field type mapping constants
    public static final String TEXT_TYPE = "text";
    public static final String DATE_TYPE = "date";
    public static final String EMAIL_TYPE = "email";
    public static final String NUMBER_TYPE = "number";
    public static final String BOOLEAN_TYPE = "boolean";
    public static final String FILE_TYPE = "file";
    public static final String SWITCH_TYPE = "switch";
    public static final String RICHTEXT_TYPE = "richtext";
    public static final String PICKONE_TYPE = "pickone";
    public static final String PICKMANY_TYPE = "pickmany";
    public static final String PHONE_TYPE = "phone";
    public static final String ADDRESS_TYPE = "address";
    public static final String CURRENCY_TYPE = "currency";
    public static final String PERCENTAGE_TYPE = "percentage";
    public static final String TIME_TYPE = "time";
    public static final String DATERANGE_TYPE = "daterange";
    public static final String DATETIMERANGE_TYPE = "datetimerange";
    public static final String LOCATION_TYPE = "location";
    public static final String FROM_KEY = "from";
    public static final String TO_KEY = "to";
    public static final String COUNTRY_CODE_KEY = "countryCode";
    public static final String NATIONAL_NUMBER_KEY = "nationalNumber";
    public static final String INCLUDE_TIME_OF_DAY_KEY = "includeTimeOfDay";
    public static final String EMPTY_STRING = "";
    public static final String DATE_RANGE_SEPARATOR = " - ";
    public static final String PHONE_NUMBER_PREFIX = "+";
    public static final String PHONE_NUMBER_SEPARATOR = " ";
    public static final String PERCENTAGE_SIGN = "%";
    public static final String PERCENTAGE_VALUE_FORMAT = "0.##";
    public static final String FORMAT_PATTERN = "formatPattern";

    // Default currency symbol and formatting
    public static final String DEFAULT_CURRENCY_SYMBOL = "$";
    public static final String DEFAULT_CURRENCY_FORMAT = "%.2f";

    // Field configuration constants
    public static final String FIELD_TYPE_KEY = "type";
    public static final String FIELD_TYPE_VALUE = "Field";
    public static final String FIELD_TYPE_ATTR_KEY = "fieldType";
    public static final String ATTRIBUTES_KEY = "attributes";
    public static final String FIELD_KEY = "field";

    // Attribute keys
    public static final String VISUAL_WIDTH_KEY = "visualWidth";
    public static final String TRIM_KEY = "trim";
    public static final String DATE_FORMAT_KEY = "dateFormat";
    public static final String SHOW_CALENDAR_KEY = "showCalendar";
    public static final String ALLOW_PAST_DATES_KEY = "allowPastDates";
    public static final String ALLOW_FUTURE_DATES_KEY = "allowFutureDates";
    public static final String MULTIPLE_FILE_UPLOAD_KEY = "multipleFileUpload";
    public static final String MIN_FILE_SIZE_UNIT_KEY = "minFileSizeUnit";
    public static final String MAX_FILE_SIZE_UNIT_KEY = "maxFileSizeUnit";
    public static final String FILE_TYPES_KEY = "fileTypes";
}