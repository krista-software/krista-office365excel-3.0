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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores;

import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.util.KeyValueStore;
import org.jvnet.hk2.annotations.Service;

import javax.inject.Inject;
import java.io.IOException;
import java.util.Map;
import java.util.UUID;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.*;

/**
 * Store for managing Office365 Excel authentication attributes.
 * This service provides methods to save, load, and remove Office365 Excel authentication
 * context information using a key-value storage mechanism.
 */
@Service
public class OfficeExcelAttributeStore {

    private final Invoker invoker;
    private final KeyValueStore store;

    @Inject
    public OfficeExcelAttributeStore(Invoker invoker, KeyValueStore store) {
        this.invoker = invoker;
        this.store = store;
    }

    /**
     * Loads Office365 Excel attributes from storage using the provided authentication context ID.
     *
     * @return An OfficeExcelAttributes object containing the loaded authentication information
     */
    public OfficeExcelAttributes load(String key) {
        final String loadedAttributes = String.valueOf(store.get(key));
        if (NULL.equals(loadedAttributes) || loadedAttributes.isBlank()) {
            return null;
        } else {
            @SuppressWarnings("unchecked")
            Map<String, Object> attributes = Constants.GSON.fromJson(loadedAttributes, Map.class);
            String authType = attributes.get(AUTH_TYPE).toString();
            String authFlow = attributes.containsKey(AUTH_FLOW) ? (String) attributes.get(AUTH_FLOW) : null;
            String baseUrl = invoker.getRoutingInfo().getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE);

            if (PUBLIC.equals(authType) && AUTH_FLOW_PRIVATE_USER.equals(authFlow)) {
                // User Authentication (own Azure AD app) — credentials must be preserved even though authType=Public
                return new OfficeExcelAttributes(
                        (String) attributes.get(CLIENT_ID),
                        (String) attributes.get(CLIENT_SECRET),
                        (String) attributes.get(TENANT_ID),
                        (String) attributes.get(EMAIL),
                        authType, AUTH_FLOW_PRIVATE_USER, baseUrl);
            } else if (PUBLIC.equals(authType)) {
                // Quick Setup — Krista-managed credentials loaded from config.properties
                return new OfficeExcelAttributes(null, null, null, (String) attributes.get(EMAIL),
                        authType, AUTH_FLOW_PUBLIC, baseUrl);
            } else if (authFlow != null) {
                // Server to Server — authFlow explicitly stored (CLIENT_CREDENTIALS_SECRET)
                return new OfficeExcelAttributes(
                        (String) attributes.get(CLIENT_ID),
                        (String) attributes.get(CLIENT_SECRET),
                        (String) attributes.get(TENANT_ID),
                        (String) attributes.get(EMAIL),
                        authType, authFlow, baseUrl);
            } else {
                // Legacy migration: old Private tab (no authFlow stored) used auth code flow with
                // the user's own Azure AD app — equivalent to the new User Authentication flow.
                // Migrate authType to PUBLIC so getCredentials() matching works with the new auth tab.
                return new OfficeExcelAttributes(
                        (String) attributes.get(CLIENT_ID),
                        (String) attributes.get(CLIENT_SECRET),
                        (String) attributes.get(TENANT_ID),
                        (String) attributes.get(EMAIL),
                        PUBLIC, AUTH_FLOW_PRIVATE_USER, baseUrl);
            }
        }
    }

    /**
     * Saves Office365 Excel attributes to storage and generates a new authentication context ID.
     *
     * @return A newly generated authentication context ID as a String
     */
    public String save(OfficeExcelAttributes attributes) {
        if (attributes == null) {
            throw new IllegalArgumentException("Office365 Excel Attributes cannot be null");
        }
        String authContextId = AUTH_CONTEXT_ID + UNDER_SCORE + UUID.randomUUID();
        Map<String, Object> attributeMap = attributes.toMap();
        store.put(authContextId, Constants.GSON.toJson(attributeMap));
        return authContextId;
    }

    public boolean save(OfficeExcelAttributes attributes, String invokerId) {
        if (attributes == null) {
            throw new IllegalArgumentException("Office365 Excel Attributes cannot be null");
        }
        Map<String, Object> attributeMap = attributes.toMap();
        store.put(invokerId, Constants.GSON.toJson(attributeMap));
        return true;
    }

    /**
     * Removes given authContext id from KeyValueStore
     *
     * @param authContextId authContext id
     */
    public void remove(String authContextId) {
        store.remove(authContextId);
    }
}
