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

import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

import static org.junit.jupiter.api.Assertions.*;

class FieldTypesTest {

    @Test
    void testTextFieldLogicalName() {
        assertEquals("com.krista.fields.Text", FieldTypes.TEXT_FIELD_LOGICAL_NAME);
    }

    @Test
    void testNumberFieldLogicalName() {
        assertEquals("com.krista.fields.Number", FieldTypes.NUMBER_FIELD_LOGICAL_NAME);
    }

    @Test
    void testSwitchFieldLogicalName() {
        assertEquals("com.krista.fields.Switch", FieldTypes.SWITCH_FIELD_LOGICAL_NAME);
    }

    @Test
    void testBooleanFieldLogicalName() {
        assertEquals("com.krista.fields.Boolean", FieldTypes.BOOLEAN_FIELD_LOGICAL_NAME);
    }

    @Test
    void testDateFieldLogicalName() {
        assertEquals("com.krista.fields.Date", FieldTypes.DATE_FIELD_LOGICAL_NAME);
    }

    @Test
    void testRichTextFieldLogicalName() {
        assertEquals("com.krista.fields.RichText", FieldTypes.RICH_TEXT_FIELD_LOGICAL_NAME);
    }

    @Test
    void testFileFieldLogicalName() {
        assertEquals("com.krista.fields.File", FieldTypes.FILE_FIELD_LOGICAL_NAME);
    }

    @Test
    void testPickOneLogicalName() {
        assertEquals("com.krista.fields.PickOne", FieldTypes.PICK_ONE_LOGICAL_NAME);
    }

    @Test
    void testListFieldLogicalName() {
        assertEquals("com.krista.fields.List", FieldTypes.LIST_FIELD_LOGICAL_NAME);
    }

    @Test
    void testCurrencyFieldLogicalName() {
        assertEquals("com.krista.fields.Currency", FieldTypes.CURRENCY_FIELD_LOGICAL_NAME);
    }

    @Test
    void testPercentageFieldLogicalName() {
        assertEquals("com.krista.fields.Percentage", FieldTypes.PERCENTAGE_FIELD_LOGICAL_NAME);
    }

    @Test
    void testPickManyLogicalName() {
        assertEquals("com.krista.fields.Checkboxes", FieldTypes.PICK_MANY_LOGICAL_NAME);
    }

    @Test
    void testAddressFieldLogicalName() {
        assertEquals("com.krista.fields.Address", FieldTypes.ADDRESS_FIELD_LOGICAL_NAME);
    }

    @Test
    void testDateRangeLogicalName() {
        assertEquals("com.krista.fields.DateRange", FieldTypes.DATE_RANGE_LOGICAL_NAME);
    }

    @Test
    void testDateTimeRangeLogicalName() {
        assertEquals("com.krista.fields.DateTimeRange", FieldTypes.DATE_TIME_RANGE_LOGICAL_NAME);
    }

    @Test
    void testPhoneFieldLogicalName() {
        assertEquals("com.krista.fields.Phone", FieldTypes.PHONE_FIELD_LOGICAL_NAME);
    }

    @Test
    void testLocationFieldLogicalName() {
        assertEquals("com.krista.fields.Location", FieldTypes.LOCATION_FIELD_LOGICAL_NAME);
    }

    @Test
    void testTimeFieldLogicalName() {
        assertEquals("com.krista.fields.Time", FieldTypes.TIME_FIELD_LOGICAL_NAME);
    }

    @Test
    void testEmailFieldLogicalName() {
        assertEquals("com.krista.fields.Email", FieldTypes.EMAIL_FIELD_LOGICAL_NAME);
    }

    @Test
    void testUnitFieldLogicalName() {
        assertEquals("com.krista.fields.Unit", FieldTypes.UNIT_FIELD_LOGICAL_NAME);
    }

    @Test
    void testAllConstantsAreNonNull() {
        assertNotNull(FieldTypes.TEXT_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.NUMBER_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.SWITCH_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.BOOLEAN_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.DATE_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.RICH_TEXT_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.FILE_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.PICK_ONE_LOGICAL_NAME);
        assertNotNull(FieldTypes.LIST_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.CURRENCY_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.PERCENTAGE_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.PICK_MANY_LOGICAL_NAME);
        assertNotNull(FieldTypes.ADDRESS_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.DATE_RANGE_LOGICAL_NAME);
        assertNotNull(FieldTypes.DATE_TIME_RANGE_LOGICAL_NAME);
        assertNotNull(FieldTypes.PHONE_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.LOCATION_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.TIME_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.EMAIL_FIELD_LOGICAL_NAME);
        assertNotNull(FieldTypes.UNIT_FIELD_LOGICAL_NAME);
    }

    @Test
    void testAllConstantsStartWithPrefix() {
        String prefix = "com.krista.fields.";
        assertTrue(FieldTypes.TEXT_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.NUMBER_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.SWITCH_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.BOOLEAN_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.DATE_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.RICH_TEXT_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.FILE_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.PICK_ONE_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.LIST_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.CURRENCY_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.PERCENTAGE_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.PICK_MANY_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.ADDRESS_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.DATE_RANGE_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.DATE_TIME_RANGE_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.PHONE_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.LOCATION_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.TIME_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.EMAIL_FIELD_LOGICAL_NAME.startsWith(prefix));
        assertTrue(FieldTypes.UNIT_FIELD_LOGICAL_NAME.startsWith(prefix));
    }

    @Test
    void testPrivateConstructor() throws Exception {
        Constructor<FieldTypes> constructor = FieldTypes.class.getDeclaredConstructor();
        assertTrue(java.lang.reflect.Modifier.isPrivate(constructor.getModifiers()));
        constructor.setAccessible(true);
        assertNotNull(constructor.newInstance());
    }
}
