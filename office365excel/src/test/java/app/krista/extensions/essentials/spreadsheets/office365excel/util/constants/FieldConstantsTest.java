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

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class FieldConstantsTest {

    @Test
    void listParsingConstants_haveCorrectValues() {
        assertEquals("[", FieldConstants.LIST_BRACKET_START);
        assertEquals("]", FieldConstants.LIST_BRACKET_END);
        assertEquals("List<", FieldConstants.LIST_ANGLE_START);
        assertEquals("List[", FieldConstants.LIST_SQUARE_START);
        assertEquals("List", FieldConstants.LIST_TYPE);
    }

    @Test
    void passwordConstants_haveCorrectValues() {
        assertEquals("Password", FieldConstants.PASSWORD_FIELD_NAME);
        assertEquals("password", FieldConstants.PASSWORD_ATTRIBUTE);
    }

    @Test
    void defaultListInnerType_isText() {
        assertEquals("Text", FieldConstants.DEFAULT_LIST_INNER_TYPE);
    }

    @Test
    void fileTypeLowercase_isFile() {
        assertEquals("file", FieldConstants.FILE_TYPE_LOWERCASE);
    }

    @Test
    void fieldTypeMappingConstants_areCorrectStrings() {
        assertEquals("text", FieldConstants.TEXT_TYPE);
        assertEquals("date", FieldConstants.DATE_TYPE);
        assertEquals("email", FieldConstants.EMAIL_TYPE);
        assertEquals("number", FieldConstants.NUMBER_TYPE);
        assertEquals("boolean", FieldConstants.BOOLEAN_TYPE);
        assertEquals("file", FieldConstants.FILE_TYPE);
        assertEquals("switch", FieldConstants.SWITCH_TYPE);
        assertEquals("richtext", FieldConstants.RICHTEXT_TYPE);
        assertEquals("pickone", FieldConstants.PICKONE_TYPE);
        assertEquals("pickmany", FieldConstants.PICKMANY_TYPE);
        assertEquals("phone", FieldConstants.PHONE_TYPE);
        assertEquals("address", FieldConstants.ADDRESS_TYPE);
        assertEquals("currency", FieldConstants.CURRENCY_TYPE);
        assertEquals("percentage", FieldConstants.PERCENTAGE_TYPE);
        assertEquals("time", FieldConstants.TIME_TYPE);
        assertEquals("daterange", FieldConstants.DATERANGE_TYPE);
        assertEquals("datetimerange", FieldConstants.DATETIMERANGE_TYPE);
        assertEquals("location", FieldConstants.LOCATION_TYPE);
    }

    @Test
    void keyConstants_haveCorrectValues() {
        assertEquals("from", FieldConstants.FROM_KEY);
        assertEquals("to", FieldConstants.TO_KEY);
        assertEquals("countryCode", FieldConstants.COUNTRY_CODE_KEY);
        assertEquals("nationalNumber", FieldConstants.NATIONAL_NUMBER_KEY);
        assertEquals("includeTimeOfDay", FieldConstants.INCLUDE_TIME_OF_DAY_KEY);
    }

    @Test
    void emptyString_isEmpty() {
        assertEquals("", FieldConstants.EMPTY_STRING);
    }

    @Test
    void formattingConstants_haveCorrectValues() {
        assertEquals(" - ", FieldConstants.DATE_RANGE_SEPARATOR);
        assertEquals("+", FieldConstants.PHONE_NUMBER_PREFIX);
        assertEquals(" ", FieldConstants.PHONE_NUMBER_SEPARATOR);
        assertEquals("%", FieldConstants.PERCENTAGE_SIGN);
        assertEquals("0.##", FieldConstants.PERCENTAGE_VALUE_FORMAT);
        assertEquals("formatPattern", FieldConstants.FORMAT_PATTERN);
    }

    @Test
    void currencyConstants_haveCorrectValues() {
        assertEquals("$", FieldConstants.DEFAULT_CURRENCY_SYMBOL);
        assertEquals("%.2f", FieldConstants.DEFAULT_CURRENCY_FORMAT);
    }

    @Test
    void fieldConfigConstants_haveCorrectValues() {
        assertEquals("type", FieldConstants.FIELD_TYPE_KEY);
        assertEquals("Field", FieldConstants.FIELD_TYPE_VALUE);
        assertEquals("fieldType", FieldConstants.FIELD_TYPE_ATTR_KEY);
        assertEquals("attributes", FieldConstants.ATTRIBUTES_KEY);
        assertEquals("field", FieldConstants.FIELD_KEY);
    }

    @Test
    void attributeKeys_haveCorrectValues() {
        assertEquals("visualWidth", FieldConstants.VISUAL_WIDTH_KEY);
        assertEquals("trim", FieldConstants.TRIM_KEY);
        assertEquals("dateFormat", FieldConstants.DATE_FORMAT_KEY);
        assertEquals("showCalendar", FieldConstants.SHOW_CALENDAR_KEY);
        assertEquals("allowPastDates", FieldConstants.ALLOW_PAST_DATES_KEY);
        assertEquals("allowFutureDates", FieldConstants.ALLOW_FUTURE_DATES_KEY);
        assertEquals("multipleFileUpload", FieldConstants.MULTIPLE_FILE_UPLOAD_KEY);
        assertEquals("minFileSizeUnit", FieldConstants.MIN_FILE_SIZE_UNIT_KEY);
        assertEquals("maxFileSizeUnit", FieldConstants.MAX_FILE_SIZE_UNIT_KEY);
        assertEquals("fileTypes", FieldConstants.FILE_TYPES_KEY);
    }

    @Test
    void allConstants_areNotNull() {
        assertNotNull(FieldConstants.LIST_BRACKET_START);
        assertNotNull(FieldConstants.LIST_BRACKET_END);
        assertNotNull(FieldConstants.LIST_ANGLE_START);
        assertNotNull(FieldConstants.LIST_SQUARE_START);
        assertNotNull(FieldConstants.LIST_TYPE);
        assertNotNull(FieldConstants.PASSWORD_FIELD_NAME);
        assertNotNull(FieldConstants.PASSWORD_ATTRIBUTE);
        assertNotNull(FieldConstants.DEFAULT_LIST_INNER_TYPE);
        assertNotNull(FieldConstants.FILE_TYPE_LOWERCASE);
        assertNotNull(FieldConstants.TEXT_TYPE);
        assertNotNull(FieldConstants.DATE_TYPE);
        assertNotNull(FieldConstants.EMAIL_TYPE);
        assertNotNull(FieldConstants.NUMBER_TYPE);
        assertNotNull(FieldConstants.BOOLEAN_TYPE);
        assertNotNull(FieldConstants.FILE_TYPE);
        assertNotNull(FieldConstants.SWITCH_TYPE);
        assertNotNull(FieldConstants.RICHTEXT_TYPE);
        assertNotNull(FieldConstants.PICKONE_TYPE);
        assertNotNull(FieldConstants.PICKMANY_TYPE);
        assertNotNull(FieldConstants.PHONE_TYPE);
        assertNotNull(FieldConstants.ADDRESS_TYPE);
        assertNotNull(FieldConstants.CURRENCY_TYPE);
        assertNotNull(FieldConstants.PERCENTAGE_TYPE);
        assertNotNull(FieldConstants.TIME_TYPE);
        assertNotNull(FieldConstants.DATERANGE_TYPE);
        assertNotNull(FieldConstants.DATETIMERANGE_TYPE);
        assertNotNull(FieldConstants.LOCATION_TYPE);
        assertNotNull(FieldConstants.FROM_KEY);
        assertNotNull(FieldConstants.TO_KEY);
        assertNotNull(FieldConstants.COUNTRY_CODE_KEY);
        assertNotNull(FieldConstants.NATIONAL_NUMBER_KEY);
        assertNotNull(FieldConstants.INCLUDE_TIME_OF_DAY_KEY);
        assertNotNull(FieldConstants.EMPTY_STRING);
        assertNotNull(FieldConstants.DATE_RANGE_SEPARATOR);
        assertNotNull(FieldConstants.PHONE_NUMBER_PREFIX);
        assertNotNull(FieldConstants.PHONE_NUMBER_SEPARATOR);
        assertNotNull(FieldConstants.PERCENTAGE_SIGN);
        assertNotNull(FieldConstants.PERCENTAGE_VALUE_FORMAT);
        assertNotNull(FieldConstants.FORMAT_PATTERN);
        assertNotNull(FieldConstants.DEFAULT_CURRENCY_SYMBOL);
        assertNotNull(FieldConstants.DEFAULT_CURRENCY_FORMAT);
        assertNotNull(FieldConstants.FIELD_TYPE_KEY);
        assertNotNull(FieldConstants.FIELD_TYPE_VALUE);
        assertNotNull(FieldConstants.FIELD_TYPE_ATTR_KEY);
        assertNotNull(FieldConstants.ATTRIBUTES_KEY);
        assertNotNull(FieldConstants.FIELD_KEY);
        assertNotNull(FieldConstants.VISUAL_WIDTH_KEY);
        assertNotNull(FieldConstants.TRIM_KEY);
        assertNotNull(FieldConstants.DATE_FORMAT_KEY);
        assertNotNull(FieldConstants.SHOW_CALENDAR_KEY);
        assertNotNull(FieldConstants.ALLOW_PAST_DATES_KEY);
        assertNotNull(FieldConstants.ALLOW_FUTURE_DATES_KEY);
        assertNotNull(FieldConstants.MULTIPLE_FILE_UPLOAD_KEY);
        assertNotNull(FieldConstants.MIN_FILE_SIZE_UNIT_KEY);
        assertNotNull(FieldConstants.MAX_FILE_SIZE_UNIT_KEY);
        assertNotNull(FieldConstants.FILE_TYPES_KEY);
    }
}
