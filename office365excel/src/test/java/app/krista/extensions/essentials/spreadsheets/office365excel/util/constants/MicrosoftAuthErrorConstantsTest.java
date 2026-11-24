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

import app.krista.extensions.essentials.spreadsheets.office365excel.util.AuthErrorRule;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class MicrosoftAuthErrorConstantsTest {

    @Test
    void authErrorRules_isNotEmpty() {
        assertNotNull(MicrosoftAuthErrorConstants.AUTH_ERROR_RULES);
        assertFalse(MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.isEmpty());
    }

    @Test
    void authErrorRules_hasExpectedSize() {
        assertEquals(9, MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.size());
    }

    @Test
    void authErrorRules_eachRuleHasNonNullUserMessage() {
        for (AuthErrorRule rule : MicrosoftAuthErrorConstants.AUTH_ERROR_RULES) {
            assertNotNull(rule.getUserMessage());
            assertFalse(rule.getUserMessage().isEmpty());
        }
    }

    @Test
    void authErrorRules_matchPasswordChangedCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.PASSWORD_CHANGED_CODE));
        assertTrue(matched, "Should match password changed error code");
    }

    @Test
    void authErrorRules_matchUserDeletedCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.USER_DELETED_CODE));
        assertTrue(matched, "Should match user deleted error code");
    }

    @Test
    void authErrorRules_matchUserDisabledCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.USER_DISABLED_CODE));
        assertTrue(matched, "Should match user disabled error code");
    }

    @Test
    void authErrorRules_matchAccountLockedCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.ACCOUNT_LOCKED_CODE));
        assertTrue(matched, "Should match account locked error code");
    }

    @Test
    void authErrorRules_matchAppNotFoundCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.APP_NOT_FOUND_CODE));
        assertTrue(matched, "Should match app not found error code");
    }

    @Test
    void authErrorRules_matchConsentRevokedCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.CONSENT_REVOKED_CODE));
        assertTrue(matched, "Should match consent revoked error code");
    }

    @Test
    void authErrorRules_matchTenantNotFoundCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.TENANT_NOT_FOUND_CODE));
        assertTrue(matched, "Should match tenant not found error code");
    }

    @Test
    void authErrorRules_matchServiceUnavailableCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.SERVICE_UNAVAILABLE_CODE));
        assertTrue(matched, "Should match service unavailable error code");
    }

    @Test
    void authErrorRules_matchInvalidClientSecretCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.INVALID_CLIENT_SECRET_CODE));
        assertTrue(matched, "Should match invalid client secret error code");
    }

    @Test
    void authErrorRules_matchRefreshTokenExpiredCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.REFRESH_TOKEN_EXPIRED_CODE));
        assertTrue(matched, "Should match refresh token expired error code");
    }

    @Test
    void authErrorRules_matchRefreshTokenRevokedCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.REFRESH_TOKEN_REVOKED_CODE));
        assertTrue(matched, "Should match refresh token revoked error code");
    }

    @Test
    void authErrorRules_matchGenericInvalidGrantCode() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches(MicrosoftAuthErrorConstants.GENERIC_INVALID_GRANT_CODE));
        assertTrue(matched, "Should match generic invalid grant error code");
    }

    @Test
    void authErrorRules_matchKeywordUserDeleted() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches("The user has been deleted from the directory"));
        assertTrue(matched, "Should match keyword 'user has been deleted'");
    }

    @Test
    void authErrorRules_matchKeywordAccountDisabled() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches("The account is disabled"));
        assertTrue(matched, "Should match keyword 'account is disabled'");
    }

    @Test
    void authErrorRules_matchKeywordServiceUnavailable() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches("service unavailable, try again later"));
        assertTrue(matched, "Should match keyword 'service unavailable'");
    }

    @Test
    void authErrorRules_matchKeywordInvalidClient() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches("invalid_client: Client authentication failed"));
        assertTrue(matched, "Should match keyword 'invalid_client'");
    }

    @Test
    void authErrorRules_matchKeywordPasswordChanged() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches("User changed or reset their password recently"));
        assertTrue(matched, "Should match keyword 'changed or reset their password'");
    }

    @Test
    void authErrorRules_passwordChangedRule_shouldRemoveToken() {
        AuthErrorRule passwordRule = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .filter(rule -> rule.matches(MicrosoftAuthErrorConstants.PASSWORD_CHANGED_CODE))
                .findFirst()
                .orElseThrow();

        assertTrue(passwordRule.shouldRemoveToken());
        assertEquals(MicrosoftAuthErrorConstants.PASSWORD_CHANGED_ERROR, passwordRule.getUserMessage());
    }

    @Test
    void authErrorRules_userDeletedRule_shouldNotRemoveToken() {
        AuthErrorRule rule = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .filter(r -> r.matches(MicrosoftAuthErrorConstants.USER_DELETED_CODE))
                .findFirst()
                .orElseThrow();

        assertFalse(rule.shouldRemoveToken());
        assertEquals(MicrosoftAuthErrorConstants.USER_DELETED_ERROR, rule.getUserMessage());
    }

    @Test
    void authErrorRules_refreshTokenExpiredRule_shouldRemoveToken() {
        AuthErrorRule rule = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .filter(r -> r.matches(MicrosoftAuthErrorConstants.REFRESH_TOKEN_EXPIRED_CODE))
                .findFirst()
                .orElseThrow();

        assertTrue(rule.shouldRemoveToken());
        assertEquals(MicrosoftAuthErrorConstants.REFRESH_TOKEN_EXPIRED_ERROR, rule.getUserMessage());
    }

    @Test
    void authErrorRules_permissionsRevokedRule_shouldRemoveToken() {
        AuthErrorRule rule = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .filter(r -> r.matches(MicrosoftAuthErrorConstants.CONSENT_REVOKED_CODE))
                .findFirst()
                .orElseThrow();

        assertTrue(rule.shouldRemoveToken());
        assertEquals(MicrosoftAuthErrorConstants.PERMISSIONS_REVOKED_ERROR, rule.getUserMessage());
    }

    @Test
    void authErrorRules_serviceUnavailableRule_shouldNotRemoveToken() {
        AuthErrorRule rule = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .filter(r -> r.matches(MicrosoftAuthErrorConstants.SERVICE_UNAVAILABLE_CODE))
                .findFirst()
                .orElseThrow();

        assertFalse(rule.shouldRemoveToken());
        assertEquals(MicrosoftAuthErrorConstants.SERVICE_UNAVAILABLE_ERROR, rule.getUserMessage());
    }

    @Test
    void authErrorRules_doesNotMatchUnrelatedMessage() {
        boolean matched = MicrosoftAuthErrorConstants.AUTH_ERROR_RULES.stream()
                .anyMatch(rule -> rule.matches("Everything is working fine"));
        assertFalse(matched, "Should not match an unrelated message");
    }

    @Test
    void errorMessageConstants_areNotNull() {
        assertNotNull(MicrosoftAuthErrorConstants.REFRESH_TOKEN_EXPIRED);
        assertNotNull(MicrosoftAuthErrorConstants.AUTHORIZATION_PROMPT);
        assertNotNull(MicrosoftAuthErrorConstants.USER_DELETED_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.USER_DISABLED_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.APP_NOT_FOUND_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.PERMISSIONS_REVOKED_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.REFRESH_TOKEN_EXPIRED_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.TENANT_NOT_FOUND_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.SERVICE_UNAVAILABLE_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.INVALID_CLIENT_SECRET_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.PASSWORD_CHANGED_ERROR);
    }

    @Test
    void errorCodeConstants_areNotNull() {
        assertNotNull(MicrosoftAuthErrorConstants.USER_DELETED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.USER_DISABLED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.ACCOUNT_LOCKED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.PASSWORD_EXPIRED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.CONSENT_REVOKED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.CONSENT_REQUIRED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.ROLE_NOT_FOUND_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.APP_NOT_FOUND_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.TENANT_NOT_FOUND_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.SERVICE_UNAVAILABLE_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.INVALID_CLIENT_SECRET_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.PASSWORD_CHANGED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.REFRESH_TOKEN_EXPIRED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.REFRESH_TOKEN_REVOKED_CODE);
        assertNotNull(MicrosoftAuthErrorConstants.PROOF_OF_POSSESSION_FAILED_CODE);
    }

    @Test
    void keywordConstants_areNotNull() {
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_USER_DELETED);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_ACCOUNT_DISABLED);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_TENANT_NOT_FOUND);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_SERVICE_UNAVAILABLE);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_NETWORK_ERROR);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_INVALID_CLIENT_SECRET);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_INVALID_CLIENT);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_INSUFFICIENT_SCOPE);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_ACCESS_DENIED);
        assertNotNull(MicrosoftAuthErrorConstants.KEYWORD_PASSWORD_CHANGED);
        assertNotNull(MicrosoftAuthErrorConstants.GENERIC_INVALID_GRANT_CODE);
    }
}
