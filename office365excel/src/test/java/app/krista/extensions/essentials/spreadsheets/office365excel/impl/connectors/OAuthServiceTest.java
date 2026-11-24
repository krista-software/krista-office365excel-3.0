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
import com.github.scribejava.core.oauth.OAuth20Service;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class OAuthServiceTest {

    @Mock
    private OfficeExcelAttributes attributes;

    private OAuthService oAuthService;

    @BeforeEach
    void setUp() {
        oAuthService = new OAuthService(attributes);
    }

    @Test
    void testGetOAuth20ServiceWithPublicAuthType() {
        when(attributes.getClientId()).thenReturn("test-client-id");
        when(attributes.getClientSecret()).thenReturn("test-client-secret");
        when(attributes.getCallbackEndPoint()).thenReturn("https://callback.example.com");
        when(attributes.getAuthType()).thenReturn("public");

        OAuth20Service service = oAuthService.getOAuth20Service();

        assertNotNull(service);
    }

    @Test
    void testGetOAuth20ServiceWithPrivateAuthType() {
        when(attributes.getClientId()).thenReturn("test-client-id");
        when(attributes.getClientSecret()).thenReturn("test-client-secret");
        when(attributes.getCallbackEndPoint()).thenReturn("https://callback.example.com");
        when(attributes.getAuthType()).thenReturn("private");
        when(attributes.getTenantId()).thenReturn("test-tenant-id");

        OAuth20Service service = oAuthService.getOAuth20Service();

        assertNotNull(service);
    }

    @Test
    void testGetOAuth20ServiceLazyInitialization() {
        when(attributes.getClientId()).thenReturn("test-client-id");
        when(attributes.getClientSecret()).thenReturn("test-client-secret");
        when(attributes.getCallbackEndPoint()).thenReturn("https://callback.example.com");
        when(attributes.getAuthType()).thenReturn("public");

        OAuth20Service first = oAuthService.getOAuth20Service();
        OAuth20Service second = oAuthService.getOAuth20Service();

        assertSame(first, second);
    }

    @Test
    void testGetOAuth20ServiceCallsGetClientId() {
        when(attributes.getClientId()).thenReturn("test-client-id");
        when(attributes.getClientSecret()).thenReturn("test-client-secret");
        when(attributes.getCallbackEndPoint()).thenReturn("https://callback.example.com");
        when(attributes.getAuthType()).thenReturn("public");

        oAuthService.getOAuth20Service();

        verify(attributes).getClientId();
        verify(attributes).getClientSecret();
        verify(attributes).getCallbackEndPoint();
        verify(attributes).getAuthType();
    }

    @Test
    void testGetOAuth20ServicePrivateCallsGetTenantId() {
        when(attributes.getClientId()).thenReturn("test-client-id");
        when(attributes.getClientSecret()).thenReturn("test-client-secret");
        when(attributes.getCallbackEndPoint()).thenReturn("https://callback.example.com");
        when(attributes.getAuthType()).thenReturn("private");
        when(attributes.getTenantId()).thenReturn("test-tenant-id");

        oAuthService.getOAuth20Service();

        verify(attributes).getTenantId();
    }
}
