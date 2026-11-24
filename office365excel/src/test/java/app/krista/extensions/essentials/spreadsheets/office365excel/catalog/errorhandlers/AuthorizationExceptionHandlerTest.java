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

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.ExtensionResponse;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.MicrosoftAuthErrorConstants.*;
import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link AuthorizationExceptionHandler}.
 */
@DisplayName("AuthorizationExceptionHandler")
class AuthorizationExceptionHandlerTest {

    @Nested
    @DisplayName("handleAuthorizationException()")
    class HandleAuthorizationException {

        @Test
        @DisplayName("requestContext=true rethrows MustAuthorizeException")
        void requestContextTrueRethrowsException() {
            MustAuthorizeException exception = new MustAuthorizeException("test error");
            assertThrows(MustAuthorizeException.class,
                    () -> AuthorizationExceptionHandler.handleAuthorizationException(exception, true));
        }

        @Test
        @DisplayName("Refresh token expired error with matching keyword returns FAILURE with user-friendly message")
        void refreshTokenExpiredError() throws MustAuthorizeException {
            // The mapping uses REFRESH_TOKEN_EXPIRED as the keyword to match, messageToDisplay = REFRESH_TOKEN_EXPIRED_ERROR
            MustAuthorizeException exception = new MustAuthorizeException("Error: " + REFRESH_TOKEN_EXPIRED);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertNotNull(response.getError());
            assertEquals(REFRESH_TOKEN_EXPIRED_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Password changed error returns appropriate response")
        void passwordChangedError() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(PASSWORD_CHANGED_ERROR);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals(PASSWORD_CHANGED_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("User deleted error returns appropriate response")
        void userDeletedError() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(USER_DELETED_ERROR);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals(USER_DELETED_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("User disabled error returns appropriate response")
        void userDisabledError() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(USER_DISABLED_ERROR);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals(USER_DISABLED_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Permissions revoked error returns appropriate response")
        void permissionsRevokedError() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(PERMISSIONS_REVOKED_ERROR);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals(PERMISSIONS_REVOKED_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("App not found error returns appropriate response")
        void appNotFoundError() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(APP_NOT_FOUND_ERROR);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals(APP_NOT_FOUND_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Tenant not found error with code returns appropriate response")
        void tenantNotFoundError() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(TENANT_NOT_FOUND_CODE);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals(TENANT_NOT_FOUND_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Tenant not found error with keyword returns appropriate response")
        void tenantNotFoundKeyword() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException("Some error: " + KEYWORD_TENANT_NOT_FOUND);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(TENANT_NOT_FOUND_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Service unavailable error with code returns appropriate response")
        void serviceUnavailableErrorCode() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(SERVICE_UNAVAILABLE_CODE);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(SERVICE_UNAVAILABLE_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Service unavailable error with keyword returns appropriate response")
        void serviceUnavailableKeyword() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException("Error: " + KEYWORD_SERVICE_UNAVAILABLE);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(SERVICE_UNAVAILABLE_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Network error keyword returns service unavailable response")
        void networkErrorKeyword() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException("Error: " + KEYWORD_NETWORK_ERROR);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(SERVICE_UNAVAILABLE_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Invalid client secret error with code returns appropriate response")
        void invalidClientSecretCode() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(INVALID_CLIENT_SECRET_CODE);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(INVALID_CLIENT_SECRET_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Invalid client secret error with keyword returns appropriate response")
        void invalidClientSecretKeyword() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException("Error: " + KEYWORD_INVALID_CLIENT_SECRET);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(INVALID_CLIENT_SECRET_ERROR, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Authorization prompt error returns appropriate response")
        void authorizationPromptError() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException(AUTHORIZATION_PROMPT);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals(AUTHORIZATION_PROMPT, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Unknown error returns generic authentication error message")
        void unknownErrorReturnsGenericMessage() throws MustAuthorizeException {
            String unknownError = "Some completely unknown error occurred";
            MustAuthorizeException exception = new MustAuthorizeException(unknownError);
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals("Authentication error: " + unknownError, response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Error response has AUTHENTICATION_ERROR exception type")
        void errorResponseHasAuthenticationType() throws MustAuthorizeException {
            MustAuthorizeException exception = new MustAuthorizeException("unknown");
            ExtensionResponse response = AuthorizationExceptionHandler.handleAuthorizationException(exception, false);
            assertEquals(ExtensionResponse.Error.ExceptionType.AUTHENTICATION_ERROR,
                    response.getError().getExceptionType());
        }
    }
}
