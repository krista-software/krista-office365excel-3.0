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

import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.extensions.util.KeyValueStore;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.Map;

@Service
public class ErrorHandlingStateManager {

    private static final Logger LOGGER = LoggerFactory.getLogger(ErrorHandlingStateManager.class);
    private final KeyValueStore entries;

    @Inject
    public ErrorHandlingStateManager(KeyValueStore internalState) {
        this.entries = internalState;

    }

    /**
     * Stores error handling state as a JSON-serialized string in the key-value store.
     * Used to preserve state across the 3-step validation flow (Confirmation → Fetch → Re-entry).
     *
     * @param key the unique identifier for the state (typically includes request ID and user context)
     * @param value the state map containing validation context and user inputs
     */
    public void put(String key, Map<String, Object> value) {
        Object keystoreValue = Constants.GSON.toJson(value);
        entries.put(key, keystoreValue);
    }

    /**
     * Retrieves and deserializes error handling state from the key-value store.
     * Converts the stored JSON string back into a map for processing.
     *
     * @param key the unique identifier for the state
     * @return map containing the deserialized state, or empty map if key not found
     */
    @SuppressWarnings("unchecked")
    public Map<String, Object> get(String key) {
        final String value = String.valueOf(entries.get(key));
        return Constants.GSON.fromJson(value, Map.class);
    }

    /**
     * Removes a key-value pair from the error handling state.
     *
     * @param key The key to remove
     * @return true if the key was removed, false if it didn't exist
     */
    public boolean remove(String key) {
        if (key == null || key.trim().isEmpty()) {
            return false;
        }

        if (entries.get(key) == null) {
            return false;
        }

        entries.remove(key);
        return true;
    }

}
