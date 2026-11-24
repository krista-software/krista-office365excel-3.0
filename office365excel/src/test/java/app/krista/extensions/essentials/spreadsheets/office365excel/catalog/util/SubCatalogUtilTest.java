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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

class SubCatalogUtilTest {

    @Test
    void testExtractFilesFromMapWithSingleFile() {
        File file = new File("test.csv");
        List<File> result = SubCatalogUtil.extractFilesFromMap(file);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals(file, result.get(0));
    }

    @Test
    void testExtractFilesFromMapWithFileList() {
        File file1 = new File("test1.csv");
        File file2 = new File("test2.csv");
        List<File> fileList = Arrays.asList(file1, file2);

        List<File> result = SubCatalogUtil.extractFilesFromMap(fileList);

        assertNotNull(result);
        assertEquals(2, result.size());
        assertEquals(file1, result.get(0));
        assertEquals(file2, result.get(1));
    }

    @Test
    void testExtractFilesFromMapWithNull() {
        List<File> result = SubCatalogUtil.extractFilesFromMap(null);

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void testExtractFilesFromMapWithNonFileObject() {
        List<File> result = SubCatalogUtil.extractFilesFromMap("not a file");

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void testExtractFilesFromMapWithEmptyList() {
        List<File> result = SubCatalogUtil.extractFilesFromMap(new ArrayList<>());

        assertNotNull(result);
        assertTrue(result.isEmpty());
    }

    @Test
    void testExtractFilesFromMapWithMixedList() {
        File file = new File("test.csv");
        List<Object> mixedList = Arrays.asList(file, "not a file", 42);

        List<File> result = SubCatalogUtil.extractFilesFromMap(mixedList);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals(file, result.get(0));
    }

    @Test
    void testProcessAndUpdateFilesWithFilesInInput() {
        File file = new File("test.csv");
        Map<String, Object> inputMap = new HashMap<>();
        inputMap.put("fileKey", file);

        Map<String, Object> stateMap = new HashMap<>();
        Map<String, Object> state = new HashMap<>();

        SubCatalogUtil.processAndUpdateFiles(inputMap, stateMap, state, "fileKey");

        assertNotNull(stateMap.get("fileKey"));
        @SuppressWarnings("unchecked")
        List<File> storedFiles = (List<File>) stateMap.get("fileKey");
        assertEquals(1, storedFiles.size());
        assertEquals(file, storedFiles.get(0));
    }

    @Test
    void testProcessAndUpdateFilesWithNoFilesInInputFallsBackToState() {
        Map<String, Object> inputMap = new HashMap<>();
        inputMap.put("fileKey", "not a file");

        Map<String, Object> stateMap = new HashMap<>();
        File stateFile = new File("state-file.csv");
        List<File> stateFiles = Collections.singletonList(stateFile);
        Map<String, Object> state = new HashMap<>();
        state.put("fileKey", stateFiles);

        SubCatalogUtil.processAndUpdateFiles(inputMap, stateMap, state, "fileKey");

        assertEquals(stateFiles, stateMap.get("fileKey"));
    }

    @Test
    void testProcessAndUpdateFilesWithNullFileObjFallsBackToState() {
        Map<String, Object> inputMap = new HashMap<>();

        Map<String, Object> stateMap = new HashMap<>();
        Map<String, Object> state = new HashMap<>();
        state.put("fileKey", null);

        SubCatalogUtil.processAndUpdateFiles(inputMap, stateMap, state, "fileKey");

        assertNull(stateMap.get("fileKey"));
    }

    @Test
    void testProcessAndUpdateFilesWithListOfFiles() {
        File file1 = new File("f1.csv");
        File file2 = new File("f2.csv");
        Map<String, Object> inputMap = new HashMap<>();
        inputMap.put("fileKey", Arrays.asList(file1, file2));

        Map<String, Object> stateMap = new HashMap<>();
        Map<String, Object> state = new HashMap<>();

        SubCatalogUtil.processAndUpdateFiles(inputMap, stateMap, state, "fileKey");

        @SuppressWarnings("unchecked")
        List<File> storedFiles = (List<File>) stateMap.get("fileKey");
        assertEquals(2, storedFiles.size());
    }
}
