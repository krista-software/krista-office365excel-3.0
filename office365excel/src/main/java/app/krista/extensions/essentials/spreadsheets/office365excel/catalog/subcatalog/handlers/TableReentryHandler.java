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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers;

import app.krista.extension.executor.ExtensionResponse;
import java.util.Map;

/**
 * Interface for table-specific reentry operations
 */
public interface TableReentryHandler {
    /**
     * Handles confirmation of table data reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of table data
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of get table rows reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleGetTableRowsConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of get table rows
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleGetTableRowsReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of vlookup in table reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleVlookupInTableConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of vlookup in table
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleVlookupInTableReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of delete column reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleDeleteColumnConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of delete column
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleDeleteColumnReentry(Map<String, Object> inputMap);
}
