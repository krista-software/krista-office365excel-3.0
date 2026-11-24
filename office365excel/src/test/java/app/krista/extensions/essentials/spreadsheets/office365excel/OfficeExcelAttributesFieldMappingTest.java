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

package app.krista.extensions.essentials.spreadsheets.office365excel;

import com.google.gson.JsonObject;
import org.junit.jupiter.api.Test;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;
import static org.junit.jupiter.api.Assertions.*;

/**
 * Test class to verify field mapping consistency with SharePoint patterns.
 * This test ensures that the JSON field names match exactly what the UI sends.
 */
public class OfficeExcelAttributesFieldMappingTest {

    private static final String BASE_URL = "http://test.com";

    @Test
    public void testCreateWithRealPayloadFromLogs() {
        // This is the exact payload from the error logs
        JsonObject payload = new JsonObject();
        payload.addProperty("authType", "Private");
        payload.addProperty("email", "simran.sethi@kristasoft.com");
        payload.addProperty("clientId", "ec0745c8-7635-4b31-97cc-d217944dd620");
        payload.addProperty("clientSecret", "REDACTED_SECRET");
        payload.addProperty("tenantId", "3694f6b4-b5f1-47ef-852f-a0b4a459ab44");
        payload.addProperty("allowMailAlert", false);
        payload.addProperty("isAllRequiredFieldsHaveValue", true);
        payload.addProperty("isCredentialsSaved", false);
        
        // This should now work without throwing an exception
        OfficeExcelAttributes attributes = OfficeExcelAttributes.create(payload, BASE_URL);
        
        assertNotNull(attributes);
        assertEquals("simran.sethi@kristasoft.com", attributes.getEmailId());
        assertEquals(PRIVATE, attributes.getAuthType());
        assertEquals("ec0745c8-7635-4b31-97cc-d217944dd620", attributes.getClientId());
        assertEquals("REDACTED_SECRET", attributes.getClientSecret());
        assertEquals("3694f6b4-b5f1-47ef-852f-a0b4a459ab44", attributes.getTenantId());
    }

    @Test
    public void testFieldMappingConsistencyWithSharePoint() {
        // Test that our constants match SharePoint patterns
        assertEquals("clientId", CLIENT_ID);  // Should be camelCase, not "clientID"
        assertEquals("clientSecret", CLIENT_SECRET);
        assertEquals("tenantId", TENANT_ID);
        assertEquals("email", EMAIL);
        assertEquals("authType", AUTH_TYPE);
    }

    @Test
    public void testPublicAuthWithCamelCaseFields() {
        JsonObject payload = new JsonObject();
        payload.addProperty("authType", "Public");
        payload.addProperty("email", "test@example.com");

        OfficeExcelAttributes attributes = OfficeExcelAttributes.create(payload, BASE_URL);

        assertNotNull(attributes);
        assertEquals("test@example.com", attributes.getEmailId());
        assertEquals(PUBLIC, attributes.getAuthType());
        // For public auth, client credentials should be loaded from config.properties
        assertNotNull(attributes.getClientId());
        assertNotNull(attributes.getClientSecret());
        assertEquals("47495500-7157-4620-bc7e-499c2ee1873e", attributes.getClientId());
        assertEquals("REDACTED_SECRET", attributes.getClientSecret());
        assertNull(attributes.getTenantId());
    }

    @Test
    public void testPrivateAuthWithAllCamelCaseFields() {
        JsonObject payload = new JsonObject();
        payload.addProperty("authType", "Private");
        payload.addProperty("email", "test@example.com");
        payload.addProperty("clientId", "test-client-id");
        payload.addProperty("clientSecret", "test-client-secret");
        payload.addProperty("tenantId", "test-tenant-id");
        
        OfficeExcelAttributes attributes = OfficeExcelAttributes.create(payload, BASE_URL);
        
        assertNotNull(attributes);
        assertEquals("test@example.com", attributes.getEmailId());
        assertEquals(PRIVATE, attributes.getAuthType());
        assertEquals("test-client-id", attributes.getClientId());
        assertEquals("test-client-secret", attributes.getClientSecret());
        assertEquals("test-tenant-id", attributes.getTenantId());
    }
}
