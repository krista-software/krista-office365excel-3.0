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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp.attributes;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.GSON;

/**
 * Utility class for pretty-printing objects using Gson.
 */
public class Print {

    private static final Logger LOGGER = LoggerFactory.getLogger(Print.class);

    /**
     * Pretty prints any object to the console along with its class name.
     *
     * @param object the object to be printed
     * @param <T>    the type of the object
     */
    public static <T> void prettyPrint(T object) {
        if (object != null) {
            try {
                String className = object.getClass().getName();
                String json = GSON.toJson(object);
                LOGGER.info("Class Name: {}\nObject Details:\n{}", className, json);
            } catch (Exception cause) {
                LOGGER.error("Error while printing object", cause);
            }
        } else {
            LOGGER.error("Object is null");
        }
    }
}
