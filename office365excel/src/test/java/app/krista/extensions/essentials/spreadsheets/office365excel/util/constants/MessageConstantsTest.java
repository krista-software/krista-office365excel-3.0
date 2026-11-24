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

class MessageConstantsTest {

    @Test
    void refreshTokenExpiredReAuthorize_isNotNull() {
        assertNotNull(MessageConstants.REFRESH_TOKEN_EXPIRED_RE_AUTHORIZE);
    }

    @Test
    void refreshTokenNotFoundReAuthorize_isNotNull() {
        assertNotNull(MessageConstants.REFRESH_TOKEN_NOT_FOUND_RE_AUTHORIZE);
    }

    @Test
    void refreshTokenExpiredReAuthorize_containsReAuthorize() {
        assertTrue(MessageConstants.REFRESH_TOKEN_EXPIRED_RE_AUTHORIZE.contains("re-authorize"));
    }

    @Test
    void refreshTokenNotFoundReAuthorize_containsAuthorize() {
        assertTrue(MessageConstants.REFRESH_TOKEN_NOT_FOUND_RE_AUTHORIZE.contains("authorize"));
    }

    @Test
    void refreshTokenExpiredReAuthorize_isNotEmpty() {
        assertFalse(MessageConstants.REFRESH_TOKEN_EXPIRED_RE_AUTHORIZE.isEmpty());
    }

    @Test
    void refreshTokenNotFoundReAuthorize_isNotEmpty() {
        assertFalse(MessageConstants.REFRESH_TOKEN_NOT_FOUND_RE_AUTHORIZE.isEmpty());
    }

    @Test
    void refreshTokenExpiredReAuthorize_hasExpectedValue() {
        assertEquals(
                "Your authentication session has expired. Please re-authorize the connection to continue.",
                MessageConstants.REFRESH_TOKEN_EXPIRED_RE_AUTHORIZE
        );
    }

    @Test
    void refreshTokenNotFoundReAuthorize_hasExpectedValue() {
        assertEquals(
                "No authentication token found. Please authorize the connection first.",
                MessageConstants.REFRESH_TOKEN_NOT_FOUND_RE_AUTHORIZE
        );
    }
}
