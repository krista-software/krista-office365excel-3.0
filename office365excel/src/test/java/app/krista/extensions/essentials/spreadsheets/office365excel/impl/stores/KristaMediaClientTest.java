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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores;

import app.krista.ksdk.files.FileHandle;
import app.krista.ksdk.files.FileRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.io.*;
import java.lang.reflect.Method;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class KristaMediaClientTest {

    @Mock
    private FileRepository fileRepository;

    @InjectMocks
    private KristaMediaClient kristaMediaClient;

    @Test
    void testToKristaFileWithSupportedFormat() throws IOException {
        File tempFile = File.createTempFile("test", ".xlsx");
        tempFile.deleteOnExit();
        Files.writeString(tempFile.toPath(), "test content");

        FileHandle mockFileHandle = mock(FileHandle.class);
        app.krista.model.base.File mockKristaFile = mock(app.krista.model.base.File.class);
        when(fileRepository.createNewFileByName(tempFile.getName())).thenReturn(mockFileHandle);
        when(mockFileHandle.getFile()).thenReturn(mockKristaFile);

        app.krista.model.base.File result = kristaMediaClient.toKristaFile(tempFile);

        assertNotNull(result);
        assertEquals(mockKristaFile, result);
        verify(fileRepository).createNewFileByName(tempFile.getName());
        verify(mockFileHandle).setContent(any(FileInputStream.class));
    }

    @Test
    void testToKristaFileWithUnsupportedFormatHtml() throws IOException {
        File tempFile = File.createTempFile("test", ".html");
        tempFile.deleteOnExit();
        Files.writeString(tempFile.toPath(), "<html>test</html>");

        FileHandle mockFileHandle = mock(FileHandle.class);
        app.krista.model.base.File mockKristaFile = mock(app.krista.model.base.File.class);

        // The zipped file will have .zip extension
        when(fileRepository.createNewFileByName(anyString())).thenReturn(mockFileHandle);
        when(mockFileHandle.getFile()).thenReturn(mockKristaFile);

        app.krista.model.base.File result = kristaMediaClient.toKristaFile(tempFile);

        assertNotNull(result);
        // Verify that the zip file was used
        verify(fileRepository).createNewFileByName(argThat(name -> name.endsWith(".zip")));
    }

    @Test
    void testToJavaFile() throws IOException {
        app.krista.model.base.File mockKristaFile = mock(app.krista.model.base.File.class);
        when(mockKristaFile.getFileName()).thenReturn("test-output.xlsx");

        FileHandle mockFileHandle = mock(FileHandle.class);
        byte[] testContent = "test content".getBytes();
        InputStream inputStream = new ByteArrayInputStream(testContent);

        when(fileRepository.getFile(mockKristaFile)).thenReturn(mockFileHandle);
        when(mockFileHandle.getContent()).thenReturn(inputStream);

        File result = kristaMediaClient.toJavaFile(mockKristaFile);

        assertNotNull(result);
        assertEquals("test-output.xlsx", result.getName());
        // Clean up
        result.delete();
    }

    @Test
    void testCompressFile() throws IOException {
        File tempFile = File.createTempFile("compress-test", ".txt");
        tempFile.deleteOnExit();
        Files.writeString(tempFile.toPath(), "content to compress");

        String zipPath = tempFile.getParent() + "/compress-test-output.zip";
        File zipFile = new File(zipPath);
        zipFile.deleteOnExit();

        KristaMediaClient.compressFile(zipPath, tempFile.getAbsolutePath());

        assertTrue(zipFile.exists());
        assertTrue(zipFile.length() > 0);
    }

    @Test
    void testGetFileExtensionViaReflection() throws Exception {
        Method method = KristaMediaClient.class.getDeclaredMethod("getFileExtension", String.class);
        method.setAccessible(true);

        String result = (String) method.invoke(kristaMediaClient, "test.xlsx");
        assertEquals("xlsx", result);
    }

    @Test
    void testGetFileExtensionWithoutDot() throws Exception {
        Method method = KristaMediaClient.class.getDeclaredMethod("getFileExtension", String.class);
        method.setAccessible(true);

        assertThrows(Exception.class, () -> method.invoke(kristaMediaClient, "testfile"));
    }

    @Test
    void testGetFileExtensionMultipleDots() throws Exception {
        Method method = KristaMediaClient.class.getDeclaredMethod("getFileExtension", String.class);
        method.setAccessible(true);

        String result = (String) method.invoke(kristaMediaClient, "test.data.csv");
        assertEquals("csv", result);
    }

    @Test
    void testIsUnsupportedFileFormatViaReflection() throws Exception {
        Method method = KristaMediaClient.class.getDeclaredMethod("isUnsupportedFileFormat", String.class);
        method.setAccessible(true);

        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.html"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.exe"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.js"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.jsp"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.php5"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.pht"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.phtml"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.shtml"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.asa"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.cer"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.asax"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.swf"));
        assertTrue((Boolean) method.invoke(kristaMediaClient, "test.xap"));
        assertFalse((Boolean) method.invoke(kristaMediaClient, "test.xlsx"));
        assertFalse((Boolean) method.invoke(kristaMediaClient, "test.csv"));
        assertFalse((Boolean) method.invoke(kristaMediaClient, "test.pdf"));
    }

    @Test
    void testToKristaFileWithJsUnsupportedFormat() throws IOException {
        File tempFile = File.createTempFile("test", ".js");
        tempFile.deleteOnExit();
        Files.writeString(tempFile.toPath(), "console.log('test');");

        FileHandle mockFileHandle = mock(FileHandle.class);
        app.krista.model.base.File mockKristaFile = mock(app.krista.model.base.File.class);

        when(fileRepository.createNewFileByName(anyString())).thenReturn(mockFileHandle);
        when(mockFileHandle.getFile()).thenReturn(mockKristaFile);

        app.krista.model.base.File result = kristaMediaClient.toKristaFile(tempFile);

        assertNotNull(result);
        verify(fileRepository).createNewFileByName(argThat(name -> name.endsWith(".zip")));
    }

    @Test
    void testToJavaFileWithEmptyContent() throws IOException {
        app.krista.model.base.File mockKristaFile = mock(app.krista.model.base.File.class);
        when(mockKristaFile.getFileName()).thenReturn("empty-file.xlsx");

        FileHandle mockFileHandle = mock(FileHandle.class);
        InputStream emptyStream = new ByteArrayInputStream(new byte[0]);

        when(fileRepository.getFile(mockKristaFile)).thenReturn(mockFileHandle);
        when(mockFileHandle.getContent()).thenReturn(emptyStream);

        File result = kristaMediaClient.toJavaFile(mockKristaFile);

        assertNotNull(result);
        assertEquals(0, result.length());
        result.delete();
    }
}
