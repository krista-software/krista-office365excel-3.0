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
 * Interface for sheet-specific reentry operations
 */
public interface SheetReentryHandler {
    /**
     * Handles confirmation of sheet data reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of sheet data
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of delete row with index reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleDeleteRowWithIndexConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of delete row with index
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleDeleteRowWithIndexReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of get complete row info reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleGetCompleteRowInfoConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of get complete row info
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleGetCompleteRowInfoReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of search reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleSearchConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of search
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleSearchReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of delete column with index reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleDeleteColumnWithIndexConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of delete column with index
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleDeleteColumnWithIndexReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of add row with index reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleAddRowWithIndexConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of add row with index
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleAddRowWithIndexReentry(Map<String, Object> inputMap);

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
     * Handles confirmation of get column values in sheet reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleGetColumnValuesInSheetConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of get column values in sheet
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleGetColumnValuesInSheetReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of clear shared sheet reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleClearSharedSheetConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of clear shared sheet
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleClearSharedSheetReentry(Map<String, Object> inputMap);
}
