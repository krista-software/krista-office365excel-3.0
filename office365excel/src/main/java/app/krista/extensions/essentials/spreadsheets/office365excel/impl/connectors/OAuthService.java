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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors;

import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.github.scribejava.apis.MicrosoftAzureActiveDirectory20Api;
import com.github.scribejava.core.builder.ServiceBuilder;
import com.github.scribejava.core.oauth.OAuth20Service;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.PUBLIC;

/**
 * OAuth service for Office365 Excel authentication.
 * 
 * <p>This service provides OAuth 2.0 authentication capabilities for Microsoft Office365 Excel
 * through the Microsoft Graph API. It supports both public and private authentication modes
 * and handles the creation and configuration of OAuth service instances.</p>
 */
@Service
public class OAuthService {

    private final OfficeExcelAttributes attributes;
    private OAuth20Service oAuth20Service;

    @Inject
    public OAuthService(OfficeExcelAttributes attributes) {
        this.attributes = attributes;
    }

    /**
     * Returns a configured OAuth 2.0 service instance for Microsoft Office365 Excel authentication with lazy initialization and thread safety.
     *
     * <p>This method implements a thread-safe lazy initialization pattern to create and cache an OAuth service instance
     * based on the configured Office Excel attributes. The service is created only once and reused for subsequent calls,
     * improving performance and ensuring consistent configuration across authentication requests.</p>
     *
     * <p>The OAuth service configuration includes:</p>
     * <ul>
     *   <li>Client ID and secret from Office Excel attributes</li>
     *   <li>Required Microsoft Graph API scopes for Excel operations</li>
     *   <li>Callback endpoint URL for OAuth flow completion</li>
     *   <li>Azure AD API instance (default or tenant-specific based on authentication type)</li>
     * </ul>
     *
     * @return A configured OAuth20Service instance ready for authorization URL generation and token exchange
     * @throws IllegalStateException if Office Excel attributes contain invalid OAuth configuration
     * @throws RuntimeException if OAuth service creation fails due to configuration or network issues
     */
    public synchronized OAuth20Service getOAuth20Service() {
        if (oAuth20Service == null) {
            oAuth20Service = new ServiceBuilder(attributes.getClientId())
                    .apiSecret(attributes.getClientSecret())
                    .defaultScope(Constants.SCOPE)
                    .callback(attributes.getCallbackEndPoint())
                    .build(attributes.getAuthType().equals(PUBLIC) ? MicrosoftAzureActiveDirectory20Api.instance()
                            : MicrosoftAzureActiveDirectory20Api.custom(attributes.getTenantId()));
        }
        return oAuth20Service;
    }
}
