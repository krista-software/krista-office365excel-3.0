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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors;

import okhttp3.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.net.HttpURLConnection;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class RestClientTest {

    private RestClient restClient;

    @BeforeEach
    void setUp() {
        restClient = new RestClient("test-access-token", null);
    }

    @Test
    void testConstructorWithDriveId() {
        RestClient client = new RestClient("token", "drive123");
        assertNotNull(client);
        assertEquals("token", client.accessToken);
        assertEquals("drive123", client.driveId);
        assertTrue(client.instanceUrl.contains("drives/drive123"));
    }

    @Test
    void testConstructorWithoutDriveId() {
        RestClient client = new RestClient("token", null);
        assertNotNull(client);
        assertTrue(client.instanceUrl.contains("/drive/items/"));
        assertFalse(client.instanceUrl.contains("drives/"));
    }

    @Test
    void testStaticConstants() {
        assertEquals("Bearer ", RestClient.BEARER);
        assertEquals("Bad Request", RestClient.BAD_REQUEST);
        assertEquals("Internal Server Error", RestClient.INTERNAL_SERVER_ERROR);
        assertEquals(" Request is Success", RestClient.REQUEST_IS_SUCCESS);
        assertEquals("Authorization", RestClient.AUTHORIZATION);
        assertEquals("Content-Type", RestClient.CONTENT_TYPE);
        assertEquals("You are not authorized", RestClient.YOU_ARE_NOT_AUTHORIZED);
        assertEquals("application/json", RestClient.APPLICATION_JSON);
        assertEquals("Status code is ", RestClient.STATUS_CODE_IS);
    }

    @Test
    void testYouAreNotAuthorized1() {
        assertNotNull(RestClient.YOU_ARE_NOT_AUTHORIZED1);
        assertTrue(RestClient.YOU_ARE_NOT_AUTHORIZED1.contains("not authorized"));
    }

    @Test
    void testOneOrMoreInputParameterIsInvalid() {
        assertNotNull(RestClient.ONE_OR_MORE_INPUT_PARAMETER_IS_INVALID);
        assertTrue(RestClient.ONE_OR_MORE_INPUT_PARAMETER_IS_INVALID.contains("invalid"));
    }

    @Test
    void testOneOrMoreFieldProvidedAreInvalid() {
        assertNotNull(RestClient.ONE_OR_MORE_FIELD_PROVIDED_ARE_INVALID);
        assertTrue(RestClient.ONE_OR_MORE_FIELD_PROVIDED_ARE_INVALID.contains("invalid"));
    }

    @Test
    void testCreateRequestBody() {
        Map<String, Object> bodyParams = new HashMap<>();
        bodyParams.put("key", "value");

        RequestBody body = RestClient.createRequestBody(bodyParams);

        assertNotNull(body);
        assertNotNull(body.contentType());
        assertTrue(body.contentType().toString().contains("application/json"));
    }

    @Test
    void testHandleResponseSuccess200() {
        Response response = createMockResponse(HttpURLConnection.HTTP_OK);

        Response result = restClient.handleResponse(response);

        assertNotNull(result);
        assertEquals(HttpURLConnection.HTTP_OK, result.code());
    }

    @Test
    void testHandleResponseCreated201() {
        Response response = createMockResponse(HttpURLConnection.HTTP_CREATED);

        Response result = restClient.handleResponse(response);

        assertNotNull(result);
        assertEquals(HttpURLConnection.HTTP_CREATED, result.code());
    }

    @Test
    void testHandleResponseNoContent204() {
        Response response = createMockResponse(HttpURLConnection.HTTP_NO_CONTENT);

        Response result = restClient.handleResponse(response);

        assertNotNull(result);
    }

    @Test
    void testHandleResponseUnauthorized401() {
        Response response = createMockResponse(HttpURLConnection.HTTP_UNAUTHORIZED);

        assertThrows(IllegalArgumentException.class, () -> restClient.handleResponse(response));
    }

    @Test
    void testHandleResponseBadRequest400() {
        Response response = createMockResponse(HttpURLConnection.HTTP_BAD_REQUEST);

        assertThrows(IllegalArgumentException.class, () -> restClient.handleResponse(response));
    }

    @Test
    void testHandleResponseInternalError500() {
        Response response = createMockResponse(HttpURLConnection.HTTP_INTERNAL_ERROR);

        assertThrows(IllegalArgumentException.class, () -> restClient.handleResponse(response));
    }

    @Test
    void testHandleResponseOtherStatusCodeWithBody() {
        Response response = createMockResponseWithBody(HttpURLConnection.HTTP_FORBIDDEN, "Forbidden");

        Response result = restClient.handleResponse(response);

        assertNull(result);
    }

    @Test
    void testHandleResponseOtherStatusCodeWithoutBody() {
        Request request = new Request.Builder().url("https://example.com").build();
        Response response = new Response.Builder()
                .request(request)
                .protocol(Protocol.HTTP_1_1)
                .code(HttpURLConnection.HTTP_FORBIDDEN)
                .message("Forbidden")
                .body(null)
                .build();

        Response result = restClient.handleResponse(response);

        assertNull(result);
    }

    @Test
    void testRequestFailedLogMessage() {
        assertNotNull(RestClient.REQUEST_FAILED_WITH_STATUS_CODE_RESPONSE_BODY);
    }

    @Test
    void testErrorReadingResponseBody() {
        assertEquals("Error reading response body", RestClient.ERROR_READING_RESPONSE_BODY);
    }

    private Response createMockResponse(int statusCode) {
        Request request = new Request.Builder().url("https://example.com").build();
        return new Response.Builder()
                .request(request)
                .protocol(Protocol.HTTP_1_1)
                .code(statusCode)
                .message("test")
                .body(ResponseBody.create("", MediaType.parse("application/json")))
                .build();
    }

    private Response createMockResponseWithBody(int statusCode, String body) {
        Request request = new Request.Builder().url("https://example.com").build();
        return new Response.Builder()
                .request(request)
                .protocol(Protocol.HTTP_1_1)
                .code(statusCode)
                .message("test")
                .body(ResponseBody.create(body, MediaType.parse("application/json")))
                .build();
    }
}
