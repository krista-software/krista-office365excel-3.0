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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers;

import app.krista.extensions.util.KeyValueStore;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Unit tests for {@link ErrorHandlingStateManager}.
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("ErrorHandlingStateManager")
class ErrorHandlingStateManagerTest {

    @Mock
    private KeyValueStore keyValueStore;

    private ErrorHandlingStateManager stateManager;

    @BeforeEach
    void setUp() {
        stateManager = new ErrorHandlingStateManager(keyValueStore);
    }

    @Nested
    @DisplayName("put()")
    class Put {

        @Test
        @DisplayName("Stores JSON-serialized map in KeyValueStore")
        void storesJsonSerializedMap() {
            Map<String, Object> value = Map.of("key1", "value1", "key2", 42);
            stateManager.put("testKey", value);
            verify(keyValueStore).put(eq("testKey"), argThat(arg -> {
                String json = (String) arg;
                return json.contains("key1") && json.contains("value1");
            }));
        }

        @Test
        @DisplayName("Stores empty map as JSON")
        void storesEmptyMap() {
            Map<String, Object> value = Map.of();
            stateManager.put("emptyKey", value);
            verify(keyValueStore).put(eq("emptyKey"), argThat(arg -> {
                String json = (String) arg;
                return json.contains("{");
            }));
        }
    }

    @Nested
    @DisplayName("get()")
    class Get {

        @Test
        @DisplayName("Returns deserialized map from KeyValueStore")
        void returnsDeserializedMap() {
            when(keyValueStore.get("testKey")).thenReturn("{\"key1\":\"value1\"}");
            Map<String, Object> result = stateManager.get("testKey");
            assertNotNull(result);
            assertEquals("value1", result.get("key1"));
        }

        @Test
        @DisplayName("Returns map for JSON with nested values")
        void returnsMapForNestedJson() {
            when(keyValueStore.get("nestedKey")).thenReturn("{\"outer\":{\"inner\":\"deep\"}}");
            Map<String, Object> result = stateManager.get("nestedKey");
            assertNotNull(result);
            assertNotNull(result.get("outer"));
        }
    }

    @Nested
    @DisplayName("remove()")
    class Remove {

        @Test
        @DisplayName("Returns true when key exists and is removed")
        void returnsTrueWhenKeyExists() {
            when(keyValueStore.get("existingKey")).thenReturn("someValue");
            boolean result = stateManager.remove("existingKey");
            assertTrue(result);
            verify(keyValueStore).remove("existingKey");
        }

        @Test
        @DisplayName("Returns false when key is null")
        void returnsFalseWhenKeyIsNull() {
            boolean result = stateManager.remove(null);
            assertFalse(result);
            verify(keyValueStore, never()).remove(any());
        }

        @Test
        @DisplayName("Returns false when key is empty string")
        void returnsFalseWhenKeyIsEmpty() {
            boolean result = stateManager.remove("");
            assertFalse(result);
            verify(keyValueStore, never()).remove(any());
        }

        @Test
        @DisplayName("Returns false when key is whitespace only")
        void returnsFalseWhenKeyIsWhitespace() {
            boolean result = stateManager.remove("   ");
            assertFalse(result);
            verify(keyValueStore, never()).remove(any());
        }

        @Test
        @DisplayName("Returns false when key does not exist in store")
        void returnsFalseWhenKeyDoesNotExist() {
            when(keyValueStore.get("nonExistingKey")).thenReturn(null);
            boolean result = stateManager.remove("nonExistingKey");
            assertFalse(result);
            verify(keyValueStore, never()).remove(any());
        }
    }
}
