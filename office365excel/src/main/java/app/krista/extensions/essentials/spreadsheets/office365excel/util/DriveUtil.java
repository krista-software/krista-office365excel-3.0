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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

public class DriveUtil {

    private DriveUtil() {
    }

    public static byte[] getByteArray(File file) throws IOException {
        if (file == null) {
            return null;
        }
        byte[] arr = new byte[(int) file.length()];
        try (FileInputStream fileInputStream = new FileInputStream(file)) {
            int count;
            do {
                count = fileInputStream.read(arr);
            } while (count > 0);
            return arr;
        }
    }

    public static File copyInputStreamToFile(InputStream inputStream, String fileId) throws IOException {
        if (inputStream == null) {
            return null;
        }
        File outputFile = new File((fileId));
        Files.copy(inputStream, outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        return outputFile;
    }
}
