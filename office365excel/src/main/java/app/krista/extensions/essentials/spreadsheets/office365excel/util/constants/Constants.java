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

import app.krista.model.base.FreeForm;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

import java.util.Map;

public final class Constants {

    private Constants() {
    }

    public static final Gson GSON = new GsonBuilder()
            .serializeSpecialFloatingPointValues()
            .setPrettyPrinting()
            .setLenient()  // Allow lenient parsing to handle malformed JSON from Microsoft Graph API
            .create();

    public static final String ENABLE_ERROR_HANDLING = "Allow Retry";

    public static final String STATE = "state";

    public static final String SPACE = " ";

    public static final String BACKWARD_SLASH = "/";

    public static final String FILES_READWRITE_ALL = "Files.ReadWrite.All";

    public static final String FILES_READWRITE = "Files.ReadWrite";

    public static final String SITES_READ_ALL = "Sites.Read.All";

    public static final String EMPTY_STRING = "";


    public static final Character A = 'A';

    public static final String SHEET_RANGE_SEPARATOR = "!";

    public static final String COLON = ":";

    public static final String COMMA = ",";

    public static final String CELL_ID = "Cell ID";

    public static final String FORMULA = "Formula";

    public static final String VALUE = "Value";

    public static final String USER_READ_WRITE = "User.ReadWrite";


    public static final String FIRST_COLUMN = "A";

    public static final String LAST_COLUMN = "XFD";

    public static final String EMAIL = "email";

    public static final String CLIENT_ID = "clientId";

    public static final String CLIENT_SECRET = "clientSecret";

    public static final String ID = "id";

    public static final String EXCEL_FILE_EXTENSION = ".xlsx";

    public static final String OPENID = "openid";

    public static final String OFFLINE_ACCESS = "offline_access";

    public static final String CODE = "code";

    public static final String JAXRS_ID = "excel365extn";

    // User-delegated scopes for public auth (no admin consent required).
    // User.Read is used for /me (Test Connection); User.ReadWrite is intentionally excluded
    // because it requires admin consent in many enterprise tenants (AADSTS65001).
    public static final String USER_READ = "User.Read";
    public static final String SCOPE = OPENID + SPACE
            + OFFLINE_ACCESS + SPACE
            + FILES_READWRITE
            + SPACE + USER_READ;

    public static final String CALLBACK = "callback";

    public static final String TENANT_ID = "tenantId";

    public static final String AUTHORITY = "https://login.microsoftonline.com/";

    public static final String SCOPE_SEPARATOR = " ";

    public static final String ORG_AUTHORITY = "https://login.microsoftonline.com/organizations/";

    public static final String MEDIA_URL = "http://platform:9991";

    public static final String MAX_VALUE = "Max_Value";

    public static final String MAX_VALUE_ROW_INDEX = "Max_Value_Row_Index";

    public static final String MIN_VALUE = "Min_Value";

    public static final String MIN_VALUE_ROW_INDEX = "Min_Value_Row_Index";

    public static final String DOWN = "Down";

    public static final String INVALID_RANGE_IS_GIVEN = "Invalid Excel range format: ";

    public static final String INVALID_COLUMN_RANGE_IS_GIVEN = "Invalid column range format: ";

    public static final String INVALID_CELL_ID_IS_GIVEN = "Invalid cell reference: ";

    public static final String WORKBOOK_WORKSHEETS = "/workbook/worksheets/";

    public static final String RANGE_ADDRESS = "/range(address='";

    public static final String SHIFT = "shift";

    public static final String ROW_DATA_NOT_PROVIDED = "No row data provided. Please provide either 'Rows' or 'Rows from CSV' to add data to the sheet.";

    public static final String DELETE = "')/delete";

    public static final String COL = ":";

    public static final String DATA = "Data";

    public static final String LEFT = "Left";

    public static final String UP = "Up";

    public static final String REFRESH_TOKEN_EXPIRED_PLEASE_REAUTHORIZE_YOURSELF = "Your authentication session has expired. Please re-authorize the connection to continue.";

    public static final String ERROR = "Error";

    public static final String TEXT = "Text";

    public static final String GRAPH_MICROSOFT_COM_V_1_0_ME = "https://graph.microsoft.com/v1.0/me";

    public static final FreeForm FREE_FORM = new FreeForm();

    public enum ApplyTo {All, Formats, Contents}

    public static final String AUTHORIZATION_PROMPT = "Please complete the authorization process and click 'Validate Attributes' before saving your configuration.";

    public static final String USER_ID = "userId";

    public static final String AUTH_CONTEXT_ID = "authContextId";

    public static final String CLIENT_USER_ID = "clientUserId";
    public static final String INVOKER_PREFIX  = "INV_";

    // Authentication constants
    public static final String WS_CONTACT = "wsContact";
    /** Suffix for per-user attrs key — avoids collision with the refresh token (bare wsContactId) in KeyValueStore. */
    public static final String PER_USER_ATTRS_SUFFIX = "_excel_attrs";
    public static final String ERROR_OCCURRED_DURING_AUTHORIZATION = "Error occurred during authorization:";
    public static final String ERROR_DESCRIPTION = "error_description";
    public static final String UNAUTHORISED_USER = "Unauthorized user:";
    public static final String USER_AUTHENTICATED_SUCCESSFULLY = "User authenticated successfully.";

    // Authentication types
    public static final String PUBLIC = "Public";
    public static final String PRIVATE = "Private";

    // Authentication flows (match SDK HTML authFlow values)
    public static final String AUTH_FLOW = "authFlow";
    public static final String AUTH_FLOW_PUBLIC = "AUTH_CODE_PUBLIC";         // Quick Setup — Krista managed
    public static final String AUTH_FLOW_PRIVATE_USER = "AUTH_CODE_PRIVATE"; // User Authentication — own Azure AD app
    public static final String AUTH_FLOW_CLIENT_CREDS = "CLIENT_CREDENTIALS_SECRET"; // Server to Server
    /** Scope for Client Credentials (app-only) token requests. */
    public static final String CLIENT_CREDENTIALS_SCOPE = "https://graph.microsoft.com/.default";

    // Authentication field names
    public static final String AUTH_TYPE = "authType";
    public static final String FAILED_TO_SAVE_ATTRIBUTES = "Unable to save your configuration. Please check your settings and try again.";
    public static final String UNSUPPORTED_AUTH = "Please select either 'Public' or 'Private' as your authentication type.";
    public static final String FAILED_TO_GET_ACCOUNT = "Failed to get account";
    public static final String GOT_ERROR_FOR_AUTHENTICATION_SO_SENDING_FOR_AUTHENTICATION = "Got error for Authentication. So sending for authentication";
    public static final String GOT_ERROR_FOR_AUTHENTICATION_SO_SENDING_FOR_RE_AUTHENTICATION = "Got error for Authentication. So sending for re authentication";
    public static final String HASH = "#";
    public static final String UNDER_SCORE = "_";
    public static final String NULL = "null";

    public static final String PUBLIC_CLIENT_SECRET = "public.clientSecret";

    public static final String PUBLIC_CLIENT_ID = "public.clientId";
    public static final String REFRESH_TOKEN_IS_NO_LONGER_VALID_PLEASE_REAUTHORIZE_YOURSELF = "Your session has expired. Please sign in again to continue accessing Office 365 Excel.";


    public static final String AUTH_URL_QUERY_PARAMS = "&access_type=offline&approval_prompt=force&prompt=select_account";
    public static final String LOCAL_EXTN_URL = "https://extension.local.eng.krista.app";
    public static final String LOCAL_EXTN_REPLACE_URL = "http://localhost:8765";
    public static final String EXTENSION_API_PATTERN = "/extension/api/.*";
    public static final String DEFAULT_CALLBACK_PATH = "/extension/api/rest/v3/oauth/callback";
    public static final String EXTENSION_FORWARD_PATH = "/rest/excel365extn/v3/oauth/callback";
    public static final String REST_EXCEL_CALLBACK = "/rest/excel365extn/callback";



    public static final String UTC = "UTC";

    public static final String MAILBOX_SETTINGS = "mailboxSettings";

    /**
     * Mapping of Windows timezone names to Java IANA timezone IDs.
     * Used to convert Microsoft Graph API timezone to Java TimeZone.
     */
    public static final Map<String, String> WINDOWS_TO_IANA = Map.ofEntries(
            Map.entry("India Standard Time", "Asia/Kolkata"),
            Map.entry("Pacific Standard Time", "America/Los_Angeles"),
            Map.entry("Mountain Standard Time", "America/Denver"),
            Map.entry("Central Standard Time", "America/Chicago"),
            Map.entry("Eastern Standard Time", "America/New_York"),
            Map.entry("US Mountain Standard Time", "America/Phoenix"),
            Map.entry("Alaskan Standard Time", "America/Anchorage"),
            Map.entry("Hawaiian Standard Time", "Pacific/Honolulu"),
            Map.entry("GMT Standard Time", "Europe/London"),
            Map.entry("W. Europe Standard Time", "Europe/Berlin"),
            Map.entry("Central Europe Standard Time", "Europe/Budapest"),
            Map.entry("Romance Standard Time", "Europe/Paris"),
            Map.entry("E. Europe Standard Time", "Europe/Bucharest"),
            Map.entry("FLE Standard Time", "Europe/Kiev"),
            Map.entry("GTB Standard Time", "Europe/Athens"),
            Map.entry("Russian Standard Time", "Europe/Moscow"),
            Map.entry("China Standard Time", "Asia/Shanghai"),
            Map.entry("Tokyo Standard Time", "Asia/Tokyo"),
            Map.entry("Korea Standard Time", "Asia/Seoul"),
            Map.entry("Singapore Standard Time", "Asia/Singapore"),
            Map.entry("W. Australia Standard Time", "Australia/Perth"),
            Map.entry("AUS Eastern Standard Time", "Australia/Sydney"),
            Map.entry("E. Australia Standard Time", "Australia/Brisbane"),
            Map.entry("West Asia Standard Time", "Asia/Tashkent"),
            Map.entry("Central Asia Standard Time", "Asia/Almaty"),
            Map.entry("SE Asia Standard Time", "Asia/Bangkok"),
            Map.entry("Myanmar Standard Time", "Asia/Rangoon"),
            Map.entry("Sri Lanka Standard Time", "Asia/Colombo"),
            Map.entry("Nepal Standard Time", "Asia/Kathmandu"),
            Map.entry("Bangladesh Standard Time", "Asia/Dhaka"),
            Map.entry("Pakistan Standard Time", "Asia/Karachi"),
            Map.entry("Arabian Standard Time", "Asia/Dubai"),
            Map.entry("Iran Standard Time", "Asia/Tehran"),
            Map.entry("Afghanistan Standard Time", "Asia/Kabul"),
            Map.entry("Israel Standard Time", "Asia/Jerusalem"),
            Map.entry("Arabic Standard Time", "Asia/Baghdad"),
            Map.entry("Turkey Standard Time", "Europe/Istanbul"),
            Map.entry("SA Pacific Standard Time", "America/Bogota"),
            Map.entry("SA Western Standard Time", "America/La_Paz"),
            Map.entry("SA Eastern Standard Time", "America/Cayenne"),
            Map.entry("Argentina Standard Time", "America/Buenos_Aires"),
            Map.entry("E. South America Standard Time", "America/Sao_Paulo"),
            Map.entry("Venezuela Standard Time", "America/Caracas"),
            Map.entry("Atlantic Standard Time", "America/Halifax"),
            Map.entry("Central America Standard Time", "America/Guatemala"),
            Map.entry("Canada Central Standard Time", "America/Regina"),
            Map.entry("Mexico Standard Time", "America/Mexico_City"),
            Map.entry("South Africa Standard Time", "Africa/Johannesburg"),
            Map.entry("Egypt Standard Time", "Africa/Cairo"),
            Map.entry("W. Central Africa Standard Time", "Africa/Lagos"),
            Map.entry("E. Africa Standard Time", "Africa/Nairobi"),
            Map.entry("New Zealand Standard Time", "Pacific/Auckland"),
            Map.entry("Fiji Standard Time", "Pacific/Fiji"),
            Map.entry("UTC", "UTC")
    );

}
