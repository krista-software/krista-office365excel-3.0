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

package app.krista.extensions.essentials.spreadsheets.office365excel;

import app.krista.sdk.msauth.extension.registry.MsAuthExtensionDescriptor;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class Office365ExcelExtensionDescriptorTest {

    private Office365ExcelExtensionDescriptor descriptor;

    @BeforeEach
    void setUp() {
        descriptor = new Office365ExcelExtensionDescriptor();
    }

    @Test
    void testDescriptorCanBeInstantiated() {
        assertNotNull(descriptor);
    }

    @Test
    void testImplementsMsAuthExtensionDescriptor() {
        assertTrue(descriptor instanceof MsAuthExtensionDescriptor);
    }

    @Test
    void testGetConfigClassReturnsOffice365ExcelAuthConfig() {
        Class<?> configClass = descriptor.getConfigClass();
        assertEquals(Office365ExcelAuthConfig.class, configClass);
    }

    @Test
    void testGetConfigClassReturnsNonNull() {
        assertNotNull(descriptor.getConfigClass());
    }

    @Test
    void testGetConfigClassReturnsSameInstanceEveryTime() {
        Class<?> first = descriptor.getConfigClass();
        Class<?> second = descriptor.getConfigClass();
        assertSame(first, second);
    }
}
