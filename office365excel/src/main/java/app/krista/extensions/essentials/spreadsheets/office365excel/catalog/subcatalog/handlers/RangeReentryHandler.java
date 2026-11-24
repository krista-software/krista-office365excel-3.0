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
 * Interface for range-specific reentry operations
 */
public interface RangeReentryHandler {
    /**
     * Handles confirmation of range data reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of range data
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of sheet vlookup reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleSheetVlookupConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of sheet vlookup
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleSheetVlookupReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of get rows for range reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleGetRowsForRangeConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of get rows for range
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleGetRowsForRangeReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of execute formula reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleExecuteFormulaConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of execute formula
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleExecuteFormulaReentry(Map<String, Object> inputMap);

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