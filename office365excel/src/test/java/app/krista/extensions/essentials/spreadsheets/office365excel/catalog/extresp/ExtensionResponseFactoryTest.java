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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp;

import app.krista.extension.executor.ExtensionResponse;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ExtensionResponseFactory}.
 */
@DisplayName("ExtensionResponseFactory")
class ExtensionResponseFactoryTest {

    @Nested
    @DisplayName("create(Map)")
    class CreateWithMap {

        @Test
        @DisplayName("Returns success response with values")
        void returnsSuccessResponseWithValues() {
            Map<String, Object> values = new HashMap<>();
            values.put("key1", "value1");
            values.put("key2", 42);

            ExtensionResponse response = ExtensionResponseFactory.create(values);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.SUCCESS, response.getResult());
            assertNotNull(response.getResponseValue());
        }

        @Test
        @DisplayName("Returns success response with empty map")
        void returnsSuccessResponseWithEmptyMap() {
            ExtensionResponse response = ExtensionResponseFactory.create(Map.of());
            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.SUCCESS, response.getResult());
        }
    }

    @Nested
    @DisplayName("create(Exception, message, type, actions, name, state)")
    class CreateWithException {

        @Test
        @DisplayName("Returns failure response with stack trace from exception")
        void returnsFailureResponseWithStackTrace() {
            Exception cause = new RuntimeException("test error");
            Map<String, Object> state = Map.of("stateKey", "stateVal");

            ExtensionResponse response = ExtensionResponseFactory.create(
                    cause,
                    "Error occurred",
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(),
                    "subCatalogName",
                    state);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertNotNull(response.getError());
            assertEquals("Error occurred", response.getError().getErrorMessage());
            assertNotNull(response.getError().getStackTrace());
            assertFalse(response.getError().getStackTrace().isEmpty());
            assertEquals(state, response.getState());
        }

        @Test
        @DisplayName("Returns failure response with null state and name")
        void returnsFailureResponseWithNullStateAndName() {
            Exception cause = new IllegalArgumentException("bad arg");

            ExtensionResponse response = ExtensionResponseFactory.create(
                    cause,
                    "Bad argument",
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(),
                    null,
                    null);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals("Bad argument", response.getError().getErrorMessage());
        }
    }

    @Nested
    @DisplayName("create(String message, type, actions, name, state)")
    class CreateWithMessage {

        @Test
        @DisplayName("Returns failure response without stack trace")
        void returnsFailureResponseWithoutStackTrace() {
            Map<String, Object> state = Map.of("key", "val");

            ExtensionResponse response = ExtensionResponseFactory.create(
                    "Something went wrong",
                    ExtensionResponse.Error.ExceptionType.AUTHENTICATION_ERROR,
                    List.of(),
                    "catalogName",
                    state);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals("Something went wrong", response.getError().getErrorMessage());
            assertEquals("", response.getError().getStackTrace());
            assertEquals(state, response.getState());
        }

        @Test
        @DisplayName("Returns failure response with null name and state")
        void returnsFailureResponseWithNullNameAndState() {
            ExtensionResponse response = ExtensionResponseFactory.create(
                    "Error message",
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(),
                    null,
                    null);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
        }
    }

    @Nested
    @DisplayName("create(message, type, actions, name, state, result)")
    class CreateWithResult {

        @Test
        @DisplayName("Returns response with specified FAILURE result")
        void returnsResponseWithFailureResult() {
            ExtensionResponse response = ExtensionResponseFactory.create(
                    "Error",
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(),
                    "catalogName",
                    null,
                    ExtensionResponse.Result.FAILURE);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.FAILURE, response.getResult());
            assertEquals("Error", response.getError().getErrorMessage());
        }

        @Test
        @DisplayName("Returns response with specified SUCCESS result")
        void returnsResponseWithSuccessResult() {
            ExtensionResponse response = ExtensionResponseFactory.create(
                    "Info message",
                    ExtensionResponse.Error.ExceptionType.LOGIC_ERROR,
                    List.of(),
                    "catalogName",
                    null,
                    ExtensionResponse.Result.SUCCESS);

            assertNotNull(response);
            assertEquals(ExtensionResponse.Result.SUCCESS, response.getResult());
        }

        @Test
        @DisplayName("Returns response with state preserved")
        void returnsResponseWithStatePreserved() {
            Map<String, Object> state = Map.of("stateKey", "stateValue");

            ExtensionResponse response = ExtensionResponseFactory.create(
                    "Msg",
                    ExtensionResponse.Error.ExceptionType.AUTHENTICATION_ERROR,
                    List.of(),
                    "name",
                    state,
                    ExtensionResponse.Result.FAILURE);

            assertEquals(state, response.getState());
        }
    }
}
