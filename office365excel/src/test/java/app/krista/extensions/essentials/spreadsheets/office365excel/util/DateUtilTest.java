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

import org.junit.BeforeClass;
import org.junit.Test;
import java.time.Instant;
import java.util.TimeZone;
import static org.junit.Assert.assertEquals;

public class DateUtilTest {

    private static final long TS = Instant.parse("2021-11-30T19:20:00Z").toEpochMilli();

    @BeforeClass
    public static void setup() {
        // Ensure stable timezone across environments
        TimeZone.setDefault(TimeZone.getTimeZone("UTC"));
    }

    @Test
    public void testFormatDate() {
        // Test case 1: Check if date formatting works with time included (UTC timezone)
        String formattedDate = DateUtil.formatDate(true, TS, "MM/dd/yyyy", "UTC");
        System.out.println("Formatted Date: " + formattedDate);
        assertEquals("11/30/2021 7:20 pm", formattedDate);

        // Test case 2: Check if date formatting works without time
        formattedDate = DateUtil.formatDate(false, TS, "MM/dd/yyyy", "UTC");
        assertEquals("11/30/2021", formattedDate);
    }

    @Test
    public void testFormatTime() {
        // Test case 1: Check if time formatting works (UTC timezone)
        String formattedTime = DateUtil.formatTime(TS, "UTC");
        System.out.println("Formatted Time: " + formattedTime);
        assertEquals("7:20 pm", formattedTime);
    }

    @Test
    public void testGetDateFormatPattern() {
        // Test pattern replacement logic
        String formattedDate = DateUtil.formatDate(false, TS, "MM/DD/YYYY", "UTC");
        // Your code replaces D->d, Y->y, so MM/DD/YYYY becomes MM/dd/yyyy
        assertEquals("11/30/2021", formattedDate);
    }

    @Test
    public void testDefaultPattern() {
        // Test with null/empty pattern - should use default
        String formattedDate = DateUtil.formatDate(false, TS, null, "UTC");
        assertEquals("Nov 30, 2021", formattedDate);

        String formattedDateWithTime = DateUtil.formatDate(true, TS, "", "UTC");
        System.out.println("Formatted Date With Time: " + formattedDateWithTime);
        assertEquals("Nov 30, 2021 7:20 pm", formattedDateWithTime);
    }

    @Test
    public void testDifferentTimezones() {
        // Test with different timezones to ensure timezone parameter works correctly

        // UTC timezone
        String utcDate = DateUtil.formatDate(true, TS, "MM/dd/yyyy", "UTC");
        assertEquals("11/30/2021 7:20 pm", utcDate);

        // Asia/Kolkata timezone (UTC+5:30)
        String istDate = DateUtil.formatDate(true, TS, "MM/dd/yyyy", "Asia/Kolkata");
        assertEquals("12/01/2021 12:50 am", istDate);

        // America/New_York timezone (UTC-5)
        String estDate = DateUtil.formatDate(true, TS, "MM/dd/yyyy", "America/New_York");
        assertEquals("11/30/2021 2:20 pm", estDate);
    }
}
