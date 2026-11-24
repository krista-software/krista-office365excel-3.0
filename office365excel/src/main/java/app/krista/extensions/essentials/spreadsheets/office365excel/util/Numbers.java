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

package app.krista.extensions.essentials.spreadsheets.office365excel.util;

import java.util.regex.Pattern;

public class Numbers {

    private Numbers() {

    }

    /**
     * This regular expression matches numbers in various formats, including integers, decimal
     * numbers, numbers with commas as thousand separators, and numbers in scientific notation.
     *
     * @return true if number
     */

    public static boolean isNumber(String input) {
        String numberPattern = "^[-+]?((\\d{1,3}(,\\d{3})*|(\\d+))(\\.\\d*)?|(\\.\\d+))([eE][-+]?\\d+)?$";
        return Pattern.matches(numberPattern, input.trim().replaceAll(",", ""));
    }

}
