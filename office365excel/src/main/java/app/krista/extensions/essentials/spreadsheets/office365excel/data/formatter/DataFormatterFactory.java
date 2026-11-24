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

package app.krista.extensions.essentials.spreadsheets.office365excel.data.formatter;

import java.util.Objects;

public class DataFormatterFactory {

    private DataFormatterFactory() {
    }

    public static DataFormatter create(String valueType, String dataFormat) {
        if (Objects.equals(valueType, "Double") && (dataFormat.contains("-") || dataFormat.contains("/"))) {
            return new DateFormatter(dataFormat);
        }
        if (Objects.equals(valueType, "Double")) {
            return new NumberFormatter(dataFormat);
        }
        return new StringFormatter(dataFormat);
    }
}
