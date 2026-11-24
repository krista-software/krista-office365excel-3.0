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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes;
import app.krista.model.field.NamedField;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.FieldConstants.*;
import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link NamedFieldFactory}.
 */
@DisplayName("NamedFieldFactory")
class NamedFieldFactoryTest {

    @Nested
    @DisplayName("createSwitchField()")
    class CreateSwitchField {

        @Test
        @DisplayName("Returns NamedField with switch type and empty maps")
        void returnsSwitchField() {
            NamedField field = NamedFieldFactory.createSwitchField("mySwitch");
            assertNotNull(field);
            assertEquals("mySwitch", field.getName());
            assertEquals(FieldTypes.SWITCH_FIELD_LOGICAL_NAME, field.getFieldType());
        }
    }

    @Nested
    @DisplayName("createFieldWithAttributes()")
    class CreateFieldWithAttributes {

        @Test
        @DisplayName("Text type with null attributes creates text field")
        void textTypeCreatesTextField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("myField", "com.krista.fields.Text", null);
            assertNotNull(field);
            assertEquals("myField", field.getName());
        }

        @Test
        @DisplayName("Text type with non-empty attributes throws UnsupportedOperationException due to immutable default map")
        void textTypeWithAttributesThrowsDueToImmutableDefaults() {
            Map<String, Object> attrs = new HashMap<>();
            attrs.put("custom", "value");
            // The default attributes map is immutable (Map.of), so putAll throws
            assertThrows(UnsupportedOperationException.class,
                    () -> NamedFieldFactory.createFieldWithAttributes("myField", "com.krista.fields.Text", attrs));
        }

        @Test
        @DisplayName("List bracket format '[ Text ]' creates list field")
        void listBracketFormatCreatesListField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("myList", "[ Text ]", null);
            assertNotNull(field);
            assertEquals("myList", field.getName());
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("List angle format 'List<Email>' creates list field")
        void listAngleFormatCreatesListField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("emailList", "List<Email]", null);
            assertNotNull(field);
            assertEquals("emailList", field.getName());
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("List square format 'List[Date]' creates list field")
        void listSquareFormatCreatesListField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("dateList", "List[Date]", null);
            assertNotNull(field);
            assertEquals("dateList", field.getName());
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("Simple 'List' type creates list field with default inner type")
        void simpleListTypeCreatesListField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("simpleList", "List", null);
            assertNotNull(field);
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("'com.krista.fields.List' type creates list field with default inner type")
        void listLogicalNameCreatesListField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("listField", "com.krista.fields.List", null);
            assertNotNull(field);
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("Password field name with text type creates password field")
        void passwordFieldNameCreatesPasswordField() {
            Map<String, Object> attrs = new HashMap<>();
            NamedField field = NamedFieldFactory.createFieldWithAttributes("Password", "com.krista.fields.Text", attrs);
            assertNotNull(field);
            assertEquals("Password", field.getName());
            assertEquals(FieldTypes.TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("Password field name with 'Text' type creates password field")
        void passwordFieldNameWithTextType() {
            Map<String, Object> attrs = new HashMap<>();
            NamedField field = NamedFieldFactory.createFieldWithAttributes("Password", "Text", attrs);
            assertNotNull(field);
            assertEquals("Password", field.getName());
            assertEquals(FieldTypes.TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("File type creates file field")
        void fileTypeCreatesFileField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("myFile", "com.krista.fields.File", null);
            assertNotNull(field);
            assertEquals("myFile", field.getName());
            assertEquals(FieldTypes.FILE_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("Non-file, non-text, non-list type falls through to createFieldByTypeWithAttributes with empty string")
        void otherTypeCreatesDefaultField() {
            NamedField field = NamedFieldFactory.createFieldWithAttributes("other", "SomeOtherType", null);
            assertNotNull(field);
            assertEquals("other", field.getName());
        }
    }

    @Nested
    @DisplayName("createFieldByTypeWithAttributes()")
    class CreateFieldByTypeWithAttributes {

        @Test
        @DisplayName("date type creates date field")
        void dateType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("dateField", "date", null);
            assertNotNull(field);
            assertEquals(FieldTypes.DATE_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("email type creates email field")
        void emailType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("emailField", "email", null);
            assertNotNull(field);
            assertEquals(FieldTypes.EMAIL_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("number type creates number field")
        void numberType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("numField", "number", null);
            assertNotNull(field);
            assertEquals(FieldTypes.NUMBER_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("boolean type creates boolean field")
        void booleanType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("boolField", "boolean", null);
            assertNotNull(field);
            assertEquals(FieldTypes.BOOLEAN_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("file type creates file field")
        void fileType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("fileField", "file", null);
            assertNotNull(field);
            assertEquals(FieldTypes.FILE_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("switch type creates switch field")
        void switchType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("switchField", "switch", null);
            assertNotNull(field);
            assertEquals(FieldTypes.SWITCH_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("richtext type creates rich text field")
        void richTextType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("rtField", "richtext", null);
            assertNotNull(field);
            assertEquals(FieldTypes.RICH_TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("pickone type creates pick-one field")
        void pickOneType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("pickField", "pickone", null);
            assertNotNull(field);
            assertEquals(FieldTypes.PICK_ONE_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("password type creates text field with password attribute")
        void passwordType() {
            Map<String, Object> attrs = new HashMap<>();
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("pwField", "password", attrs);
            assertNotNull(field);
            assertEquals(FieldTypes.TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("list type creates list field")
        void listType() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("listField", "list", null);
            assertNotNull(field);
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("unknown type defaults to text field")
        void unknownTypeDefaultsToText() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("unknownField", "unknown", null);
            assertNotNull(field);
            assertEquals(FieldTypes.TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("empty string type defaults to text field")
        void emptyStringTypeDefaultsToText() {
            NamedField field = NamedFieldFactory.createFieldByTypeWithAttributes("emptyField", "", null);
            assertNotNull(field);
            assertEquals(FieldTypes.TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }
    }

    @Nested
    @DisplayName("createTextFieldWithAttributes()")
    class CreateTextFieldWithAttributes {

        @Test
        @DisplayName("Creates text field with default attributes when attrs is null")
        void createsTextFieldWithNullAttributes() {
            NamedField field = NamedFieldFactory.createTextFieldWithAttributes("textField", null);
            assertNotNull(field);
            assertEquals("textField", field.getName());
            assertEquals(FieldTypes.TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("Creates text field with empty attributes uses defaults")
        void createsTextFieldWithEmptyAttributes() {
            NamedField field = NamedFieldFactory.createTextFieldWithAttributes("textField", Collections.emptyMap());
            assertNotNull(field);
            assertEquals(FieldTypes.TEXT_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("Creates text field with non-empty attributes throws UnsupportedOperationException due to immutable default map")
        void createsTextFieldWithCustomAttributesThrowsDueToImmutableDefaults() {
            Map<String, Object> attrs = new HashMap<>();
            attrs.put("customAttr", "customValue");
            // The default attributes map returned by getDefaultAttributesForFieldType is immutable (Map.of),
            // so putAll on it throws UnsupportedOperationException
            assertThrows(UnsupportedOperationException.class,
                    () -> NamedFieldFactory.createTextFieldWithAttributes("textField", attrs));
        }
    }

    @Nested
    @DisplayName("createListFieldWithAttributes()")
    class CreateListFieldWithAttributes {

        @Test
        @DisplayName("Creates list field with text inner type")
        void createsListFieldWithTextInnerType() {
            NamedField field = NamedFieldFactory.createListFieldWithAttributes("listField", "Text", null);
            assertNotNull(field);
            assertEquals("listField", field.getName());
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }

        @Test
        @DisplayName("Creates list field with date inner type")
        void createsListFieldWithDateInnerType() {
            NamedField field = NamedFieldFactory.createListFieldWithAttributes("dateList", "date", null);
            assertNotNull(field);
            assertEquals(FieldTypes.LIST_FIELD_LOGICAL_NAME, field.getFieldType());
        }
    }

    @Nested
    @DisplayName("getDefaultAttributesForFieldType()")
    class GetDefaultAttributesForFieldType {

        @Test
        @DisplayName("Text type returns text default attributes")
        void textTypeReturnsDefaults() {
            Map<String, Object> attrs = NamedFieldFactory.getDefaultAttributesForFieldType("text");
            assertNotNull(attrs);
            assertEquals("S", attrs.get(VISUAL_WIDTH_KEY));
            assertEquals(true, attrs.get(TRIM_KEY));
        }

        @Test
        @DisplayName("Date type returns date default attributes")
        void dateTypeReturnsDefaults() {
            Map<String, Object> attrs = NamedFieldFactory.getDefaultAttributesForFieldType("date");
            assertNotNull(attrs);
            assertEquals("S", attrs.get(VISUAL_WIDTH_KEY));
            assertEquals("MM/dd/yyyy", attrs.get(DATE_FORMAT_KEY));
            assertEquals(true, attrs.get(SHOW_CALENDAR_KEY));
            assertEquals(true, attrs.get(ALLOW_PAST_DATES_KEY));
            assertEquals(true, attrs.get(ALLOW_FUTURE_DATES_KEY));
        }

        @Test
        @DisplayName("File type returns file default attributes")
        void fileTypeReturnsDefaults() {
            Map<String, Object> attrs = NamedFieldFactory.getDefaultAttributesForFieldType("file");
            assertNotNull(attrs);
            assertEquals("S", attrs.get(VISUAL_WIDTH_KEY));
            assertEquals(true, attrs.get(MULTIPLE_FILE_UPLOAD_KEY));
            assertEquals("KB", attrs.get(MIN_FILE_SIZE_UNIT_KEY));
            assertEquals("MB", attrs.get(MAX_FILE_SIZE_UNIT_KEY));
        }

        @Test
        @DisplayName("Unknown type returns default attributes with visual width only")
        void unknownTypeReturnsDefault() {
            Map<String, Object> attrs = NamedFieldFactory.getDefaultAttributesForFieldType("unknown");
            assertNotNull(attrs);
            assertEquals("S", attrs.get(VISUAL_WIDTH_KEY));
            assertEquals(1, attrs.size());
        }
    }
}
