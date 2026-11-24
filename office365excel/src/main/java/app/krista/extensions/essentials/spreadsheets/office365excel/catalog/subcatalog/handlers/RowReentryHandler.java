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
 * Handler for row-related reentry operations
 */
public interface RowReentryHandler {

    /**
     * Handles confirmation of update row with index reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleUpdateRowWithIndexConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of update row with index
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleUpdateRowWithIndexReentry(Map<String, Object> inputMap);
    /**
     * Handles the confirmation for getting complete row information
     *
     * @param map The input map containing reentry parameters
     * @return The extension response
     */
    ExtensionResponse handleGetCompleteRowInfoConfirmation(Map<String, Object> map);

    /**
     * Handles the reentry for getting complete row information
     *
     * @param map The input map containing reentry parameters
     * @return The extension response
     */
    ExtensionResponse handleGetCompleteRowInfoReentry(Map<String, Object> map);

    /**
     * Handles confirmation of delete row with index reentry
     *
     * @param map Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleDeleteRowWithIndexConfirmation(Map<String, Object> map);

    /**
     * Processes the reentry of delete row with index
     *
     * @param map Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleDeleteRowWithIndexReentry(Map<String, Object> map);

    /**
     * Handles confirmation of add rows reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleAddRowsConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of add rows
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleAddRowsReentry(Map<String, Object> inputMap);
}