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

public class FieldTypes {

    private FieldTypes(){}

    // Basic field types
    public static final String TEXT_FIELD_LOGICAL_NAME = "com.krista.fields.Text";
    public static final String NUMBER_FIELD_LOGICAL_NAME = "com.krista.fields.Number";
    public static final String SWITCH_FIELD_LOGICAL_NAME = "com.krista.fields.Switch";
    public static final String BOOLEAN_FIELD_LOGICAL_NAME = "com.krista.fields.Boolean";
    public static final String DATE_FIELD_LOGICAL_NAME = "com.krista.fields.Date";
    public static final String RICH_TEXT_FIELD_LOGICAL_NAME = "com.krista.fields.RichText";
    public static final String FILE_FIELD_LOGICAL_NAME = "com.krista.fields.File";
    public static final String PICK_ONE_LOGICAL_NAME = "com.krista.fields.PickOne";

    // Additional field types from reference
    public static final String LIST_FIELD_LOGICAL_NAME = "com.krista.fields.List";
    public static final String CURRENCY_FIELD_LOGICAL_NAME = "com.krista.fields.Currency";
    public static final String PERCENTAGE_FIELD_LOGICAL_NAME = "com.krista.fields.Percentage";
    public static final String PICK_MANY_LOGICAL_NAME = "com.krista.fields.Checkboxes";
    public static final String ADDRESS_FIELD_LOGICAL_NAME = "com.krista.fields.Address";
    public static final String DATE_RANGE_LOGICAL_NAME = "com.krista.fields.DateRange";
    public static final String DATE_TIME_RANGE_LOGICAL_NAME = "com.krista.fields.DateTimeRange";
    public static final String PHONE_FIELD_LOGICAL_NAME = "com.krista.fields.Phone";
    public static final String LOCATION_FIELD_LOGICAL_NAME = "com.krista.fields.Location";
    public static final String TIME_FIELD_LOGICAL_NAME = "com.krista.fields.Time";
    public static final String EMAIL_FIELD_LOGICAL_NAME = "com.krista.fields.Email";
    public static final String UNIT_FIELD_LOGICAL_NAME = "com.krista.fields.Unit";
}
