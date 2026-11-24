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

/**
 * Response object for authentication operations.
 * 
 * <p>This class represents the response structure for authentication-related operations
 * such as testing connections, saving credentials, and OAuth callbacks. It provides
 * a standardized format for communicating authentication results to the client.</p>
 */
public class AuthenticationResponse {
    private final boolean isSuccess;
    private final String errorMessage;
    private final String url;

    /**
     * Constructs an AuthenticationResponse with the specified parameters.
     *
     * @param success      Whether the authentication operation was successful
     * @param errorMessage Error or informational message, if any
     * @param url          Authorization URL for OAuth flow, if applicable
     */
    public AuthenticationResponse(boolean success, String errorMessage, String url) {
        this.isSuccess = success;
        this.errorMessage = errorMessage;
        this.url = url;
    }

    /**
     * Returns whether the authentication operation was successful.
     *
     * @return true if successful, false otherwise
     */
    public boolean isSuccess() {
        return isSuccess;
    }

    /**
     * Returns the error message associated with the authentication response.
     *
     * @return Error or informational message, or null if none
     */
    public String getErrorMessage() {
        return errorMessage;
    }

    /**
     * Returns the authorization URL for OAuth flow.
     *
     * @return Authorization URL, or null if not applicable
     */
    public String getUrl() {
        return url;
    }
}
