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

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class AuthErrorRuleTest {

    @Test
    void matches_messageContainsIndicator_returnsTrue() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS50020", "user has been deleted"),
                "User deleted error",
                false
        );

        assertTrue(rule.matches("Error: AADSTS50020 occurred during login"));
    }

    @Test
    void matches_messageContainsSecondIndicator_returnsTrue() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS50020", "user has been deleted"),
                "User deleted error",
                false
        );

        assertTrue(rule.matches("The user has been deleted from the directory"));
    }

    @Test
    void matches_noIndicatorsMatch_returnsFalse() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS50020", "user has been deleted"),
                "User deleted error",
                false
        );

        assertFalse(rule.matches("Some completely unrelated error message"));
    }

    @Test
    void matches_emptyMessage_returnsFalse() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS50020"),
                "Error message",
                true
        );

        assertFalse(rule.matches(""));
    }

    @Test
    void matches_exactIndicatorMatch_returnsTrue() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("invalid_grant"),
                "Token expired",
                true
        );

        assertTrue(rule.matches("invalid_grant"));
    }

    @Test
    void getUserMessage_returnsConfiguredMessage() {
        String expectedMessage = "Your session has expired. Please sign in again.";
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS700082"),
                expectedMessage,
                true
        );

        assertEquals(expectedMessage, rule.getUserMessage());
    }

    @Test
    void shouldRemoveToken_returnsTrueWhenConfigured() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS700082"),
                "Token expired",
                true
        );

        assertTrue(rule.shouldRemoveToken());
    }

    @Test
    void shouldRemoveToken_returnsFalseWhenConfigured() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS50020"),
                "User deleted",
                false
        );

        assertFalse(rule.shouldRemoveToken());
    }

    @Test
    void matches_multipleIndicators_matchesAny() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("code1", "code2", "code3"),
                "Error",
                false
        );

        assertTrue(rule.matches("message with code3 in it"));
        assertFalse(rule.matches("message with code4 in it"));
    }

    @Test
    void matches_singleIndicator_works() {
        AuthErrorRule rule = new AuthErrorRule(
                List.of("AADSTS700016"),
                "App not found",
                false
        );

        assertTrue(rule.matches("AADSTS700016: Application not found"));
        assertFalse(rule.matches("AADSTS50020: User deleted"));
    }
}
