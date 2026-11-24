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
 * Interface for cell-specific reentry operations
 */
public interface CellReentryHandler {
    /**
     * Handles confirmation of cell data reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleAddDataToCellConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of cell data
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleAddDataToCellReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of cell data retrieval reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleCellDataConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of cell data retrieval
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleCellDataReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of add multiple cell data reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleAddMultipleCellDataConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of add multiple cell data
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleAddMultipleCellDataReentry(Map<String, Object> inputMap);

    /**
     * Handles confirmation of add multiple cell data reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleAddMultipleCellDataConfirmationForSheet(Map<String, Object> inputMap);

    /**
     * Processes the reentry of add multiple cell data
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleAddMultipleCellDataReentryForSheet(Map<String, Object> inputMap);


    /**
     * Handles confirmation of get cell value reentry
     * @param inputMap Map containing reentry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleGetCellValueConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the reentry of get cell value
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleGetCellValueReentry(Map<String, Object> inputMap);
    /**
     * Handles confirmation of add data to cell retry
     * @param inputMap Map containing retry parameters
     * @return Extension response based on user's choice
     */
    ExtensionResponse handleAddDataToCellRetryConfirmation(Map<String, Object> inputMap);

    /**
     * Processes the retry of add data to cell
     * @param inputMap Map containing corrected data
     * @return Extension response with results or error information
     */
    ExtensionResponse handleAddDataToCellRetry(Map<String, Object> inputMap);
}
