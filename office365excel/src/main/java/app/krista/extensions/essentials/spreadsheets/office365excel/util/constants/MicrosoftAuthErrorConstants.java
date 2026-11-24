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

import java.util.List;

public class MicrosoftAuthErrorConstants {

    public static final String REFRESH_TOKEN_EXPIRED = "Your session has expired. Please sign in again to continue working with Excel files.";
    public static final String AUTHORIZATION_PROMPT = "Authentication required. Click 'Validate Attributes' to sign in with your Microsoft account before saving.";

    // Error messages for ExtensionResponse
    public static final String USER_DELETED_ERROR = "Your Microsoft account is no longer active. Contact your IT administrator to restore access, or sign in with a different account.";
    public static final String USER_DISABLED_ERROR = "Your Microsoft account has been disabled. Contact your IT administrator to unlock your account. If you recently changed your password, try signing in again.";
    public static final String APP_NOT_FOUND_ERROR = "We couldn’t find the application in your Microsoft setup. Please contact your administrator.";
    public static final String PERMISSIONS_REVOKED_ERROR = "Your access permissions have been revoked or the application consent has been withdrawn. Please contact your system administrator to restore the necessary permissions.";
    public static final String REFRESH_TOKEN_EXPIRED_ERROR = "Your authentication session has expired. Please re-authorize the connection or contact your system administrator if this problem persists.";
    public static final String TENANT_NOT_FOUND_ERROR = "We couldn’t find your Microsoft organization. Please contact your administrator.";
    public static final String SERVICE_UNAVAILABLE_ERROR = "Microsoft services are temporarily unavailable. Please wait a few minutes and try again. If the issue persists, contact your system administrator.";
    public static final String INVALID_CLIENT_SECRET_ERROR = "Something’s wrong with the application’s connection. Please contact your administrator.";
    public static final String PASSWORD_CHANGED_ERROR = "Your Microsoft password was recently changed. Please re-authorize the connection to continue.";

    // Error codes
    public static final String USER_DELETED_CODE = "AADSTS50020";
    public static final String USER_DISABLED_CODE = "AADSTS50057";
    public static final String ACCOUNT_LOCKED_CODE = "AADSTS50053";
    public static final String PASSWORD_EXPIRED_CODE = "AADSTS50055";
    public static final String CONSENT_REVOKED_CODE = "AADSTS65001";
    public static final String CONSENT_REQUIRED_CODE = "AADSTS70019";
    public static final String ROLE_NOT_FOUND_CODE = "AADSTS90094";
    public static final String APP_NOT_FOUND_CODE = "AADSTS700016";
    public static final String TENANT_NOT_FOUND_CODE = "AADSTS90002";
    public static final String SERVICE_UNAVAILABLE_CODE = "AADSTS50000";
    public static final String INVALID_CLIENT_SECRET_CODE = "AADSTS7000215";
    public static final String PASSWORD_CHANGED_CODE = "AADSTS50173";
    public static final String REFRESH_TOKEN_EXPIRED_CODE = "AADSTS700082";
    public static final String REFRESH_TOKEN_REVOKED_CODE = "AADSTS700084";
    public static final String PROOF_OF_POSSESSION_FAILED_CODE = "AADSTS54005";

    // Authentication error keywords for detection
    public static final String KEYWORD_USER_DELETED = "user has been deleted";
    public static final String KEYWORD_ACCOUNT_DISABLED = "account is disabled";
    public static final String KEYWORD_TENANT_NOT_FOUND = "tenant not found";
    public static final String KEYWORD_SERVICE_UNAVAILABLE = "service unavailable";
    public static final String KEYWORD_NETWORK_ERROR = "network error";
    public static final String KEYWORD_INVALID_CLIENT_SECRET = "invalid client secret";
    public static final String KEYWORD_INVALID_CLIENT = "invalid_client";
    public static final String KEYWORD_INSUFFICIENT_SCOPE = "insufficient_scope";
    public static final String KEYWORD_ACCESS_DENIED = "access_denied";
    public static final String KEYWORD_PASSWORD_CHANGED = "changed or reset their password";
    public static final String GENERIC_INVALID_GRANT_CODE = "invalid_grant";

    public static final List<AuthErrorRule> AUTH_ERROR_RULES = List.of(
            new AuthErrorRule(List.of(PASSWORD_CHANGED_CODE, KEYWORD_PASSWORD_CHANGED), PASSWORD_CHANGED_ERROR, true),
            new AuthErrorRule(List.of(USER_DELETED_CODE, KEYWORD_USER_DELETED), USER_DELETED_ERROR, false),
            new AuthErrorRule(List.of(USER_DISABLED_CODE, ACCOUNT_LOCKED_CODE, PASSWORD_EXPIRED_CODE, KEYWORD_ACCOUNT_DISABLED), USER_DISABLED_ERROR, false),
            new AuthErrorRule(List.of(APP_NOT_FOUND_CODE), APP_NOT_FOUND_ERROR, false),
            new AuthErrorRule(List.of(CONSENT_REVOKED_CODE, CONSENT_REQUIRED_CODE, ROLE_NOT_FOUND_CODE, KEYWORD_INSUFFICIENT_SCOPE, KEYWORD_ACCESS_DENIED), PERMISSIONS_REVOKED_ERROR, true),
            new AuthErrorRule(List.of(TENANT_NOT_FOUND_CODE, KEYWORD_TENANT_NOT_FOUND), TENANT_NOT_FOUND_ERROR, false),
            new AuthErrorRule(List.of(SERVICE_UNAVAILABLE_CODE, KEYWORD_SERVICE_UNAVAILABLE, KEYWORD_NETWORK_ERROR), SERVICE_UNAVAILABLE_ERROR, false),
            new AuthErrorRule(List.of(INVALID_CLIENT_SECRET_CODE, KEYWORD_INVALID_CLIENT_SECRET, KEYWORD_INVALID_CLIENT), INVALID_CLIENT_SECRET_ERROR, false),
            new AuthErrorRule(List.of(REFRESH_TOKEN_EXPIRED_CODE, REFRESH_TOKEN_REVOKED_CODE, PROOF_OF_POSSESSION_FAILED_CODE, GENERIC_INVALID_GRANT_CODE), REFRESH_TOKEN_EXPIRED_ERROR, true)
    );
}

