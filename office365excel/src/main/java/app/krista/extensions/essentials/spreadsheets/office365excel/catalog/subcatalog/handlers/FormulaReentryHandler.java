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
 * Handler for formula-related reentry operations
 */
public interface FormulaReentryHandler {
    
    /**
     * Handles confirmation of execute formula retry
     * @param inputMap Map containing retry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleExecuteFormulaRetryConfirmation(Map<String, Object> inputMap);
    
    /**
     * Processes the retry of execute formula
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleExecuteFormulaRetry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of add formula to cell reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleAddFormulaToCellConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of add formula to cell
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleAddFormulaToCellReentry(Map<String, Object> inputMap);
}