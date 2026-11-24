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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class DriveTest {

    @Test
    void constructor_setsAllFields_fieldsAccessible() {
        Drive drive = new Drive("drive-123", "John Doe", "Jane Doe");

        assertEquals("drive-123", drive.driveId);
        assertEquals("John Doe", drive.sharedBy);
        assertEquals("Jane Doe", drive.createdBy);
    }

    @Test
    void constructor_withNullValues_fieldsAreNull() {
        Drive drive = new Drive(null, null, null);

        assertNull(drive.driveId);
        assertNull(drive.sharedBy);
        assertNull(drive.createdBy);
    }

    @Test
    void constructor_withEmptyStrings_fieldsAreEmpty() {
        Drive drive = new Drive("", "", "");

        assertEquals("", drive.driveId);
        assertEquals("", drive.sharedBy);
        assertEquals("", drive.createdBy);
    }

    @Test
    void fields_arePubliclyAccessible_canBeModified() {
        Drive drive = new Drive("id1", "shared1", "created1");

        drive.driveId = "id2";
        drive.sharedBy = "shared2";
        drive.createdBy = "created2";

        assertEquals("id2", drive.driveId);
        assertEquals("shared2", drive.sharedBy);
        assertEquals("created2", drive.createdBy);
    }

    @Test
    void constructor_withMixedNullAndValues_setsCorrectly() {
        Drive drive = new Drive("drive-456", null, "Creator");

        assertEquals("drive-456", drive.driveId);
        assertNull(drive.sharedBy);
        assertEquals("Creator", drive.createdBy);
    }
}
