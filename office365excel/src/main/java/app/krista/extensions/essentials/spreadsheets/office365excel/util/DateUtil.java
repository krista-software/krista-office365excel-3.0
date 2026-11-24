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

import java.sql.Time;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.TimeZone;

public class DateUtil {

    /**
     * Formats a date value with optional time component using the specified timezone.
     *
     * @param isTimeIncluded whether to include time in the formatted output
     * @param dateValue      the date value in milliseconds since epoch
     * @param dateFormat     the date format pattern (can be null for default format)
     * @param timezone       the timezone to use for formatting (e.g., "America/New_York", "UTC", "Asia/Kolkata").
     *                       If null or empty, uses the system default timezone.
     * @return formatted date string
     */
    public static String formatDate(boolean isTimeIncluded, Long dateValue, String dateFormat, String timezone) {
        Date date = new Date(dateValue);

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(getDateFormatPattern(isTimeIncluded, dateFormat), Locale.ENGLISH);
        if (timezone != null && !timezone.trim().isEmpty()) {
            simpleDateFormat.setTimeZone(TimeZone.getTimeZone(timezone));
        }

        String formattedDate = simpleDateFormat.format(date);

        // Ensure AM/PM is lowercase for consistency across environments
        if (isTimeIncluded) {
            formattedDate = formattedDate.replace(" AM", " am").replace(" PM", " pm");
        }

        return formattedDate;
    }

    private static String getDateFormatPattern(boolean isTimeIncluded, String dateFormat) {

        if (dateFormat != null && !dateFormat.isEmpty()) {
            String replacedDateFormat = dateFormat.replace("D", "d")
                    .replace("Y", "y")
                    .replace("o", "d");
            return isTimeIncluded ? replacedDateFormat + " h:mm a" : replacedDateFormat;
        } else {
            return isTimeIncluded ? "MMM d, yyyy h:mm a" : "MMM d, yyyy";
        }
    }

    /**
     * Formats a time value using the specified timezone.
     *
     * @param timeValue the time value in milliseconds since epoch
     * @param timezone  the timezone to use for formatting (e.g., "America/New_York", "UTC", "Asia/Kolkata").
     *                  If null or empty, uses the system default timezone.
     * @return formatted time string
     */
    public static String formatTime(Long timeValue, String timezone) {
        Time time = new Time(timeValue);
        SimpleDateFormat timeFormat = new SimpleDateFormat("h:mm a", Locale.ENGLISH);
        if (timezone != null && !timezone.trim().isEmpty()) {
            timeFormat.setTimeZone(TimeZone.getTimeZone(timezone));
        }
        String formattedTime = timeFormat.format(time);

        // Ensure AM/PM is lowercase for consistency across environments
        return formattedTime.replace(" AM", " am").replace(" PM", " pm");
    }

    private DateUtil() {

    }

}
