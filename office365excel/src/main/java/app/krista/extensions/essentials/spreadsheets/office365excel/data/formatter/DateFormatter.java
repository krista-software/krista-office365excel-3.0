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

import java.text.ParseException;

public class DateFormatter implements DataFormatter {
    //TODO: format will be used when commented line 16,17 will be in place.
    private final String format;

    public DateFormatter(String format) {
        this.format = format.replace("m", "M");
    }

    @Override
    public Object format(String data) throws ParseException {
        // TODO: below lines will be uncommented after story "https://antbrains.atlassian.net/browse/KR-7432" gets completed.
        /*SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
        return simpleDateFormat.parse(data);*/
        return data;
    }
}
