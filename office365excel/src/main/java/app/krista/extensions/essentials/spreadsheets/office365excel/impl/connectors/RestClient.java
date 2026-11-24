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

import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import okhttp3.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.util.Map;

public class RestClient {

    public static final String BEARER = "Bearer ";
    public static final String BAD_REQUEST = "Bad Request";
    public static final String INTERNAL_SERVER_ERROR = "Internal Server Error";
    public static final String REQUEST_IS_SUCCESS = " Request is Success";
    public static final String YOU_ARE_NOT_AUTHORIZED1 = "You are not authorized to access this Excel file. Please verify that you have the necessary permissions or contact your system administrator.";
    public static final String ONE_OR_MORE_INPUT_PARAMETER_IS_INVALID = "One or more input parameters are invalid. Please check that all required fields are provided with correct values and try again.";
    public static final String ONE_OR_MORE_FIELD_PROVIDED_ARE_INVALID = "One or more fields contain invalid data. Please verify your input values match the expected format and try again.";
    public static final String REQUEST_FAILED_WITH_STATUS_CODE_RESPONSE_BODY = "Request failed with status code {}. Response body: {}";
    public static final String ERROR_READING_RESPONSE_BODY = "Error reading response body";
    String accessToken;
    String driveId;
    String instanceUrl;

    public static final String AUTHORIZATION = "Authorization";

    public static final String CONTENT_TYPE = "Content-Type";

    public static final String YOU_ARE_NOT_AUTHORIZED = "You are not authorized";

    public static final String APPLICATION_JSON = "application/json";

    public static final String STATUS_CODE_IS = "Status code is ";

    private static final Logger LOGGER = LoggerFactory.getLogger(RestClient.class);

    public RestClient(String accessToken, String driveId) {
        this.accessToken = accessToken;
        this.driveId = driveId;
        this.instanceUrl = getProperInstanceUrl(driveId);
    }

    private String getProperInstanceUrl(String driveId) {
        return driveId != null ? Constants.GRAPH_MICROSOFT_COM_V_1_0_ME + "/drives/" + driveId + "/" + "items" + "/"
                : Constants.GRAPH_MICROSOFT_COM_V_1_0_ME + "/drive/items/";
    }

    public Response post(String endPoint, Map<String, Object> bodyParams) {
        try {
            OkHttpClient client = new OkHttpClient();
            RequestBody requestBody = createRequestBody(bodyParams);
            LOGGER.info("{}{}", instanceUrl, endPoint);
            Request request = new Request.Builder()
                    .url(instanceUrl + endPoint)
                    .addHeader(AUTHORIZATION, BEARER + accessToken)
                    .addHeader(CONTENT_TYPE, APPLICATION_JSON)
                    .post(requestBody)
                    .build();
            return handleResponse(client.newCall(request).execute());

        } catch (IOException cause) {
            LOGGER.debug(cause.getMessage(), cause.getCause());
            return null;
        }
    }

    public Response get(String endPoint) {
        try {
            OkHttpClient client = new OkHttpClient();
            Request request = new Request.Builder()
                    .url(instanceUrl + endPoint)
                    .addHeader(AUTHORIZATION, BEARER + accessToken)
                    .get()
                    .build();
            return handleResponse(client.newCall(request).execute());

        } catch (IOException cause) {
            LOGGER.debug(cause.getMessage(), cause.getCause());
            return null;
        }
    }


    public static RequestBody createRequestBody(Map<String, Object> bodyParams) {
        String json = Constants.GSON.toJson(bodyParams);
        MediaType mediaType = MediaType.parse(APPLICATION_JSON);
        return RequestBody.create(json, mediaType);
    }

    public Response handleResponse(Response response) {
        try {
            int statusCode = response.code();
            if (statusCode == HttpURLConnection.HTTP_OK || statusCode == HttpURLConnection.HTTP_CREATED || statusCode == HttpURLConnection.HTTP_NO_CONTENT) {
                LOGGER.info(STATUS_CODE_IS + "{}" + REQUEST_IS_SUCCESS, response.code());
                return response;
            } else if (statusCode == HttpURLConnection.HTTP_UNAUTHORIZED) {
                LOGGER.info(STATUS_CODE_IS + "{}" + YOU_ARE_NOT_AUTHORIZED, response.code());
                throw new IllegalArgumentException(YOU_ARE_NOT_AUTHORIZED1);
            } else if (statusCode == HttpURLConnection.HTTP_BAD_REQUEST) {
                LOGGER.info(STATUS_CODE_IS + "{}" + BAD_REQUEST, response.code());
                throw new IllegalArgumentException(ONE_OR_MORE_INPUT_PARAMETER_IS_INVALID);
            } else if (statusCode == HttpURLConnection.HTTP_INTERNAL_ERROR) {
                LOGGER.info(STATUS_CODE_IS + "{}" + INTERNAL_SERVER_ERROR, response.code());
                throw new IllegalArgumentException(ONE_OR_MORE_FIELD_PROVIDED_ARE_INVALID);
            } else {
                if (response.body() != null) {
                    LOGGER.error(REQUEST_FAILED_WITH_STATUS_CODE_RESPONSE_BODY, statusCode, response.body().string());
                }
                LOGGER.error(REQUEST_FAILED_WITH_STATUS_CODE_RESPONSE_BODY, statusCode, "No response body");
            }
        } catch (IOException cause) {
            LOGGER.error(ERROR_READING_RESPONSE_BODY, cause);
        }
        return null;
    }
}




