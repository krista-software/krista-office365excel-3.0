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

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class AuthenticationResponseTest {

    @Test
    void constructor_setsAllThreeFields_correctly() {
        AuthenticationResponse response = new AuthenticationResponse(true, "Success", "https://auth.example.com");

        assertTrue(response.isSuccess());
        assertEquals("Success", response.getErrorMessage());
        assertEquals("https://auth.example.com", response.getUrl());
    }

    @Test
    void constructor_failureResponse_isSuccessFalse() {
        AuthenticationResponse response = new AuthenticationResponse(false, "Invalid credentials", null);

        assertFalse(response.isSuccess());
        assertEquals("Invalid credentials", response.getErrorMessage());
        assertNull(response.getUrl());
    }

    @Test
    void constructor_withNullErrorMessage_returnsNull() {
        AuthenticationResponse response = new AuthenticationResponse(true, null, "https://example.com");

        assertTrue(response.isSuccess());
        assertNull(response.getErrorMessage());
        assertEquals("https://example.com", response.getUrl());
    }

    @Test
    void constructor_withNullUrl_returnsNull() {
        AuthenticationResponse response = new AuthenticationResponse(false, "Error occurred", null);

        assertFalse(response.isSuccess());
        assertEquals("Error occurred", response.getErrorMessage());
        assertNull(response.getUrl());
    }

    @Test
    void constructor_withAllNullStrings_successTrue() {
        AuthenticationResponse response = new AuthenticationResponse(true, null, null);

        assertTrue(response.isSuccess());
        assertNull(response.getErrorMessage());
        assertNull(response.getUrl());
    }

    @Test
    void constructor_withAllNullStrings_successFalse() {
        AuthenticationResponse response = new AuthenticationResponse(false, null, null);

        assertFalse(response.isSuccess());
        assertNull(response.getErrorMessage());
        assertNull(response.getUrl());
    }

    @Test
    void isSuccess_returnsTrue_whenConstructedWithTrue() {
        AuthenticationResponse response = new AuthenticationResponse(true, "msg", "url");
        assertTrue(response.isSuccess());
    }

    @Test
    void isSuccess_returnsFalse_whenConstructedWithFalse() {
        AuthenticationResponse response = new AuthenticationResponse(false, "msg", "url");
        assertFalse(response.isSuccess());
    }

    @Test
    void getErrorMessage_returnsEmptyString_whenSetEmpty() {
        AuthenticationResponse response = new AuthenticationResponse(true, "", "url");
        assertEquals("", response.getErrorMessage());
    }

    @Test
    void getUrl_returnsEmptyString_whenSetEmpty() {
        AuthenticationResponse response = new AuthenticationResponse(true, "msg", "");
        assertEquals("", response.getUrl());
    }
}
