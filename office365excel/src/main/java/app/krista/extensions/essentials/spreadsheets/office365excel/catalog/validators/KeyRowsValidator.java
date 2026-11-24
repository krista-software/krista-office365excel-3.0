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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.KEY_ROWS;

public class KeyRowsValidator implements Validator {
    private static final Logger LOGGER = LoggerFactory.getLogger(KeyRowsValidator.class);

    @Override
    public Boolean validate(String keyRowsStr, Map<ValidationResource, String> context) {
        if (keyRowsStr == null || keyRowsStr.trim().isEmpty()) {
            LOGGER.debug("Key rows validation failed: value is null or empty");
            return false;
        }
        return true;
    }

    @Override
    public String getConfirmStepMessage(String value, Map<ValidationResource, String> context) {
        return "The key rows provided are invalid. Would you like to re-enter them?";
    }

    @Override
    public String getFetchStepMessage() {
        return "Please provide valid key rows.";
    }

    @Override
    public String getErrMessage(String value) {
        return "Key rows cannot be empty.";
    }

    @Override
    public String getFetchFieldName() {
        return KEY_ROWS;
    }

    @Override
    public String getFieldType() {
        return "[ Text ]";
    }
}