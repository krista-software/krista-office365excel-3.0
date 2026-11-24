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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores;

import app.krista.extensions.util.KeyValueStore;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class RefreshTokenStoreTest {

    @Mock
    private KeyValueStore keyValueStore;

    private RefreshTokenStore refreshTokenStore;

    @BeforeEach
    void setUp() {
        refreshTokenStore = new RefreshTokenStore(keyValueStore);
    }

    @Test
    void testPutDelegatesToKeyValueStore() throws IOException {
        refreshTokenStore.put("tokenKey", "refreshToken123");

        verify(keyValueStore).put("tokenKey", "refreshToken123");
    }

    @Test
    void testGetDelegatesToKeyValueStore() throws IOException {
        when(keyValueStore.get("tokenKey", String.class)).thenReturn("refreshToken123");

        String result = refreshTokenStore.get("tokenKey");

        assertEquals("refreshToken123", result);
        verify(keyValueStore).get("tokenKey", String.class);
    }

    @Test
    void testGetReturnsNullWhenNotFound() throws IOException {
        when(keyValueStore.get("missingKey", String.class)).thenReturn(null);

        String result = refreshTokenStore.get("missingKey");

        assertNull(result);
    }

    @Test
    void testRemoveDelegatesToKeyValueStore() {
        refreshTokenStore.remove("tokenKey");

        verify(keyValueStore).remove("tokenKey");
    }

    @Test
    void testPutWithNullKey() throws IOException {
        refreshTokenStore.put(null, "token");

        verify(keyValueStore).put(null, "token");
    }

    @Test
    void testPutWithNullValue() throws IOException {
        refreshTokenStore.put("key", null);

        verify(keyValueStore).put("key", null);
    }

    @Test
    void testGetWithNullKey() throws IOException {
        when(keyValueStore.get(null, String.class)).thenReturn(null);

        String result = refreshTokenStore.get(null);

        assertNull(result);
    }

    @Test
    void testRemoveWithNullKey() {
        refreshTokenStore.remove(null);

        verify(keyValueStore).remove(null);
    }

    @Test
    void testPutThrowsRuntimeException() {
        doThrow(new RuntimeException("Store error")).when(keyValueStore).put("key", "token");

        assertThrows(RuntimeException.class, () -> refreshTokenStore.put("key", "token"));
    }

    @Test
    void testGetThrowsRuntimeException() {
        when(keyValueStore.get("key", String.class)).thenThrow(new RuntimeException("Read error"));

        assertThrows(RuntimeException.class, () -> refreshTokenStore.get("key"));
    }
}
