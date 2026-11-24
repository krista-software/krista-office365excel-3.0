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

package app.krista.extensions.essentials.spreadsheets.office365excel.util.constants;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ConstantsTest {

    @Test
    void gson_isNotNull() {
        assertNotNull(Constants.GSON);
    }

    @Test
    void scope_containsOpenid() {
        assertTrue(Constants.SCOPE.contains(Constants.OPENID));
    }

    @Test
    void scope_containsOfflineAccess() {
        assertTrue(Constants.SCOPE.contains(Constants.OFFLINE_ACCESS));
    }

    @Test
    void scope_containsFilesReadWrite() {
        assertTrue(Constants.SCOPE.contains(Constants.FILES_READWRITE));
    }

    @Test
    void scope_containsUserRead() {
        assertTrue(Constants.SCOPE.contains(Constants.USER_READ));
    }

    @Test
    void scope_isCorrectlyFormatted() {
        assertEquals("openid offline_access Files.ReadWrite User.Read", Constants.SCOPE);
    }

    @Test
    void authority_isCorrectUrl() {
        assertEquals("https://login.microsoftonline.com/", Constants.AUTHORITY);
    }

    @Test
    void orgAuthority_isCorrectUrl() {
        assertEquals("https://login.microsoftonline.com/organizations/", Constants.ORG_AUTHORITY);
    }

    @Test
    void keyStringConstants_areNotNull() {
        assertNotNull(Constants.STATE);
        assertNotNull(Constants.SPACE);
        assertNotNull(Constants.BACKWARD_SLASH);
        assertNotNull(Constants.FILES_READWRITE_ALL);
        assertNotNull(Constants.FILES_READWRITE);
        assertNotNull(Constants.SITES_READ_ALL);
        assertNotNull(Constants.EMPTY_STRING);
        assertNotNull(Constants.SHEET_RANGE_SEPARATOR);
        assertNotNull(Constants.COLON);
        assertNotNull(Constants.COMMA);
        assertNotNull(Constants.CELL_ID);
        assertNotNull(Constants.FORMULA);
        assertNotNull(Constants.VALUE);
        assertNotNull(Constants.EMAIL);
        assertNotNull(Constants.CLIENT_ID);
        assertNotNull(Constants.CLIENT_SECRET);
        assertNotNull(Constants.ID);
        assertNotNull(Constants.EXCEL_FILE_EXTENSION);
        assertNotNull(Constants.CODE);
        assertNotNull(Constants.JAXRS_ID);
        assertNotNull(Constants.CALLBACK);
        assertNotNull(Constants.TENANT_ID);
        assertNotNull(Constants.SCOPE_SEPARATOR);
        assertNotNull(Constants.MEDIA_URL);
        assertNotNull(Constants.MAX_VALUE);
        assertNotNull(Constants.MAX_VALUE_ROW_INDEX);
        assertNotNull(Constants.MIN_VALUE);
        assertNotNull(Constants.MIN_VALUE_ROW_INDEX);
    }

    @Test
    void characterA_isUppercaseA() {
        assertEquals('A', Constants.A);
    }

    @Test
    void emptyString_isEmpty() {
        assertEquals("", Constants.EMPTY_STRING);
    }

    @Test
    void excelFileExtension_isXlsx() {
        assertEquals(".xlsx", Constants.EXCEL_FILE_EXTENSION);
    }

    @Test
    void jaxrsId_hasExpectedValue() {
        assertEquals("excel365extn", Constants.JAXRS_ID);
    }

    @Test
    void maxValueConstants_haveExpectedValues() {
        assertEquals("Max_Value", Constants.MAX_VALUE);
        assertEquals("Max_Value_Row_Index", Constants.MAX_VALUE_ROW_INDEX);
        assertEquals("Min_Value", Constants.MIN_VALUE);
        assertEquals("Min_Value_Row_Index", Constants.MIN_VALUE_ROW_INDEX);
    }

    @Test
    void graphMicrosoftUrl_isCorrect() {
        assertEquals("https://graph.microsoft.com/v1.0/me", Constants.GRAPH_MICROSOFT_COM_V_1_0_ME);
    }

    @Test
    void freeForm_isNotNull() {
        assertNotNull(Constants.FREE_FORM);
    }

    @Test
    void authConstants_areNotNull() {
        assertNotNull(Constants.USER_ID);
        assertNotNull(Constants.AUTH_CONTEXT_ID);
        assertNotNull(Constants.CLIENT_USER_ID);
        assertNotNull(Constants.INVOKER_PREFIX);
        assertNotNull(Constants.WS_CONTACT);
        assertNotNull(Constants.PER_USER_ATTRS_SUFFIX);
        assertNotNull(Constants.ERROR_OCCURRED_DURING_AUTHORIZATION);
        assertNotNull(Constants.ERROR_DESCRIPTION);
        assertNotNull(Constants.UNAUTHORISED_USER);
        assertNotNull(Constants.USER_AUTHENTICATED_SUCCESSFULLY);
    }

    @Test
    void authFlowConstants_haveExpectedValues() {
        assertEquals("Public", Constants.PUBLIC);
        assertEquals("Private", Constants.PRIVATE);
        assertEquals("authFlow", Constants.AUTH_FLOW);
        assertEquals("AUTH_CODE_PUBLIC", Constants.AUTH_FLOW_PUBLIC);
        assertEquals("AUTH_CODE_PRIVATE", Constants.AUTH_FLOW_PRIVATE_USER);
        assertEquals("CLIENT_CREDENTIALS_SECRET", Constants.AUTH_FLOW_CLIENT_CREDS);
    }

    @Test
    void clientCredentialsScope_isCorrect() {
        assertEquals("https://graph.microsoft.com/.default", Constants.CLIENT_CREDENTIALS_SCOPE);
    }

    @Test
    void windowsToIana_isNotEmpty() {
        assertFalse(Constants.WINDOWS_TO_IANA.isEmpty());
    }

    @Test
    void windowsToIana_containsExpectedMappings() {
        assertEquals("Asia/Kolkata", Constants.WINDOWS_TO_IANA.get("India Standard Time"));
        assertEquals("America/Los_Angeles", Constants.WINDOWS_TO_IANA.get("Pacific Standard Time"));
        assertEquals("UTC", Constants.WINDOWS_TO_IANA.get("UTC"));
    }

    @Test
    void workbookConstants_areCorrect() {
        assertEquals("/workbook/worksheets/", Constants.WORKBOOK_WORKSHEETS);
        assertEquals("/range(address='", Constants.RANGE_ADDRESS);
        assertEquals("')/delete", Constants.DELETE);
    }

    @Test
    void applyToEnum_hasAllValues() {
        assertEquals(3, Constants.ApplyTo.values().length);
        assertNotNull(Constants.ApplyTo.All);
        assertNotNull(Constants.ApplyTo.Formats);
        assertNotNull(Constants.ApplyTo.Contents);
    }

    @Test
    void callbackPaths_areNotNull() {
        assertNotNull(Constants.DEFAULT_CALLBACK_PATH);
        assertNotNull(Constants.EXTENSION_FORWARD_PATH);
        assertNotNull(Constants.REST_EXCEL_CALLBACK);
        assertNotNull(Constants.EXTENSION_API_PATTERN);
    }
}
