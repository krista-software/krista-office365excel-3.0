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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes;
import app.krista.model.base.EntityValue;
import app.krista.model.entity.EntityAttributeField;
import app.krista.model.entity.EntityDefinition;
import org.junit.Test;

import java.util.*;

import static org.junit.Assert.*;
import static org.mockito.Mockito.*;

public class DataModificationTest {

    @Test
    public void testFromListOfEntities() {
        // Create mock EntityDefinition
        EntityDefinition mockEntityDefinition = mock(EntityDefinition.class);

        // Create mock EntityAttributeField for each field in test data
        EntityAttributeField mockNameField = mock(EntityAttributeField.class);
        when(mockNameField.getFieldType()).thenReturn(FieldTypes.TEXT_FIELD_LOGICAL_NAME);
        when(mockNameField.getAttributeValues()).thenReturn(new HashMap<>());

        EntityAttributeField mockValueField = mock(EntityAttributeField.class);
        when(mockValueField.getFieldType()).thenReturn(FieldTypes.NUMBER_FIELD_LOGICAL_NAME);
        when(mockValueField.getAttributeValues()).thenReturn(new HashMap<>());

        // Create map of attribute fields
        Map<String, EntityAttributeField> attributeFields = new HashMap<>();
        attributeFields.put("name", mockNameField);
        attributeFields.put("value", mockValueField);

        when(mockEntityDefinition.getAttributeFields()).thenReturn(attributeFields);

        // Create test EntityValue objects
        Map<String, Object> fields1 = new HashMap<>();
        fields1.put("name", "Entity1");
        fields1.put("value", 100);

        List<EntityValue> entityList = new ArrayList<>();
        entityList.add(new EntityValue("Entity1", fields1));

        // Call the method
        List<Map<String, Object>> result = DataModification.fromListOfEntities(entityList, null, mockEntityDefinition);

        // Assertions
        assertEquals(entityList.size(), result.size());
    }


    @Test
    public void testConvertToListOfLists_EmptyListOfMap() {
        List<Map<String, Object>> listOfMap = Collections.emptyList();
        List<List<String>> result = DataModification.convertToListOfLists(listOfMap);

        assertEquals(1, result.size());
    }

    @Test
    public void testConvertToListOfLists_SingleMap() {
        Map<String, Object> map = new HashMap<>();
        map.put("key1", "value1");
        map.put("key2", 123);

        List<Map<String, Object>> listOfMap = Collections.singletonList(map);
        List<List<String>> result = DataModification.convertToListOfLists(listOfMap);

        assertEquals(2, result.size());
        assertEquals(Arrays.asList("key1", "key2"), result.get(0));
        assertEquals(Arrays.asList("value1", "123"), result.get(1));
    }

    @Test
    public void testConvertToListOfLists_MultipleMaps() {
        Map<String, Object> map1 = new HashMap<>();
        map1.put("key1", "value1");
        map1.put("key2", 123);

        Map<String, Object> map2 = new HashMap<>();
        map2.put("key1", "value2");
        map2.put("key2", 456);

        List<Map<String, Object>> listOfMap = new ArrayList<>();
        listOfMap.add(map1);
        listOfMap.add(map2);

        List<List<String>> result = DataModification.convertToListOfLists(listOfMap);

        assertEquals(3, result.size());
        assertEquals(Arrays.asList("key1", "key2"), result.get(0));
        assertEquals(Arrays.asList("value1", "123"), result.get(1));
        assertEquals(Arrays.asList("value2", "456"), result.get(2));
    }


    @Test
    public void testGetProperHeader_MultipleMaps() {
        Map<String, Object> map1 = new HashMap<>();
        map1.put("key1", "value1");
        map1.put("key2", 123);

        Map<String, Object> map2 = new HashMap<>();
        map2.put("key1", "value2");
        map2.put("key2", 456);
        map2.put("key3", true);

        List<Map<String, Object>> dataList = new ArrayList<>();
        dataList.add(map1);
        dataList.add(map2);

        List<String> result = DataModification.getProperHeader(dataList);

        assertEquals(3, result.size());
        assertTrue(result.contains("key1"));
        assertTrue(result.contains("key2"));
        assertTrue(result.contains("key3"));
    }


    @Test
    public void testGetProperHeader_EmptyDataList() {
        List<Map<String, Object>> dataList = Collections.emptyList();
        List<String> result = DataModification.getProperHeader(dataList);

        assertEquals(0, result.size());
    }

    @Test
    public void testGetProperHeader_SingleMap() {
        Map<String, Object> map = new HashMap<>();
        map.put("key1", "value1");
        map.put("key2", 123);

        List<Map<String, Object>> dataList = Collections.singletonList(map);
        List<String> result = DataModification.getProperHeader(dataList);

        assertEquals(2, result.size());
        assertTrue(result.contains("key1"));
        assertTrue(result.contains("key2"));
    }


    @Test
    public void testConvertToListOfListsFromIndex_ValidIndex() {
        List<List<String>> listOfLists = new ArrayList<>();
        listOfLists.add(Arrays.asList("A", "B", "C"));
        listOfLists.add(Arrays.asList("1", "2", "3"));
        int startIndex = 1;
        List<List<String>> result = DataModification.convertToListOfListsFromIndex(listOfLists, startIndex);

        assertEquals(1, result.size());
        assertEquals(Arrays.asList("1", "2", "3"), result.getFirst());
    }

    @Test
    public void testConvertToListOfListsFromIndex_InvalidIndex() {
        List<List<String>> listOfLists = new ArrayList<>();
        listOfLists.add(Arrays.asList("A", "B", "C"));
        int startIndex = 3;
        List<List<String>> result = DataModification.convertToListOfListsFromIndex(listOfLists, startIndex);

        assertEquals(0, result.size());
    }

}
