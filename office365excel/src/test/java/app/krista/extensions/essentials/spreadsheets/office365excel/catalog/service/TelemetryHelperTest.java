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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service;

import app.krista.ksdk.telemetry.TelemetryMetrics;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class TelemetryHelperTest {

    @Mock
    private TelemetryMetrics telemetry;

    private TelemetryHelper telemetryHelper;

    @BeforeEach
    void setUp() {
        telemetryHelper = new TelemetryHelper(telemetry);
    }

    @Test
    void testRecordSuccessIncrementsCounterAndRecordsDuration() {
        Map<String, String> tags = new HashMap<>();
        tags.put("key1", "val1");
        long startTime = System.currentTimeMillis() - 100;

        telemetryHelper.recordSuccess("metric.test", startTime, tags);

        verify(telemetry).incrementCounter(eq("metric.test.success"), eq(1L));
        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "success".equals(map.get("status")) && "val1".equals(map.get("key1"))
        ));
    }

    @Test
    void testRecordSuccessWithNullTags() {
        long startTime = System.currentTimeMillis() - 50;

        telemetryHelper.recordSuccess("metric.test", startTime, null);

        verify(telemetry).incrementCounter(eq("metric.test.success"), eq(1L));
        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "success".equals(map.get("status"))
        ));
    }

    @Test
    void testRecordSuccessWithNullValueInTags() {
        Map<String, String> tags = new HashMap<>();
        tags.put("key1", null);
        long startTime = System.currentTimeMillis() - 50;

        telemetryHelper.recordSuccess("metric.test", startTime, tags);

        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "NA".equals(map.get("key1"))
        ));
    }

    @Test
    void testRecordRetryPrompted() {
        Map<String, String> tags = new HashMap<>();
        tags.put("op", "test");
        long startTime = System.currentTimeMillis() - 100;

        telemetryHelper.recordRetryPrompted("metric.test", startTime, tags);

        verify(telemetry).incrementCounter(eq("metric.test.retry_prompted"), eq(1L));
        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "retry_prompted".equals(map.get("status"))
        ));
    }

    @Test
    void testRecordValidationError() {
        Map<String, String> tags = new HashMap<>();
        long startTime = System.currentTimeMillis() - 100;

        telemetryHelper.recordValidationError("metric.test", startTime, "Invalid input", tags);

        verify(telemetry).incrementCounter(eq("metric.test.mustAuthorizeException"), eq(1L));
        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "validation_error".equals(map.get("status")) &&
                        "validation".equals(map.get("error_type")) &&
                        "Invalid input".equals(map.get("error_message"))
        ));
    }

    @Test
    void testRecordValidationErrorWithNullMessage() {
        Map<String, String> tags = new HashMap<>();
        long startTime = System.currentTimeMillis() - 50;

        telemetryHelper.recordValidationError("metric.test", startTime, null, tags);

        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "Unknown error".equals(map.get("error_message"))
        ));
    }

    @Test
    void testRecordError() {
        Map<String, String> tags = new HashMap<>();
        long startTime = System.currentTimeMillis() - 100;
        Exception ex = new RuntimeException("Something went wrong");

        telemetryHelper.recordError("metric.test", startTime, ex, tags);

        verify(telemetry).incrementCounter(eq("metric.test.error"), eq(1L));
        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "error".equals(map.get("status")) &&
                        "general".equals(map.get("error_type")) &&
                        "Something went wrong".equals(map.get("error_message")) &&
                        "RuntimeException".equals(map.get("error_class"))
        ));
    }

    @Test
    void testRecordErrorWithNullMessage() {
        Map<String, String> tags = new HashMap<>();
        long startTime = System.currentTimeMillis() - 50;
        Exception ex = new RuntimeException((String) null);

        telemetryHelper.recordError("metric.test", startTime, ex, tags);

        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "Unknown error".equals(map.get("error_message"))
        ));
    }

    @Test
    void testRecordErrorWithNullTags() {
        long startTime = System.currentTimeMillis() - 50;
        Exception ex = new IllegalArgumentException("bad input");

        telemetryHelper.recordError("metric.test", startTime, ex, null);

        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "error".equals(map.get("status")) &&
                        "IllegalArgumentException".equals(map.get("error_class"))
        ));
    }

    @Test
    void testIncrementCount() {
        telemetryHelper.incrementCount("metric.test");

        verify(telemetry).incrementCounter(eq("metric.test.count"), eq(1L));
    }

    @Test
    void testSafeTagMap() {
        Map<String, String> result = TelemetryHelper.safeTagMap("key1", "val1", "key2", "val2");

        assertEquals("val1", result.get("key1"));
        assertEquals("val2", result.get("key2"));
    }

    @Test
    void testSafeTagMapWithNullValue() {
        Map<String, String> result = TelemetryHelper.safeTagMap("key1", null);

        assertEquals("NA", result.get("key1"));
    }

    @Test
    void testSafeTagMapWithEmptyArgs() {
        Map<String, String> result = TelemetryHelper.safeTagMap();
        assertTrue(result.isEmpty());
    }

    @Test
    void testSafeTagMapWithOddArgs() {
        Map<String, String> result = TelemetryHelper.safeTagMap("key1", "val1", "key2");
        assertEquals(1, result.size());
        assertEquals("val1", result.get("key1"));
    }

    @Test
    void testRecordCounter() {
        Map<String, String> tags = Map.of("t1", "v1");
        telemetryHelper.recordCounter("metric.counter", 5, tags);

        verify(telemetry).incrementCounter("metric.counter", 5, tags);
    }

    @Test
    void testRecordGauge() {
        telemetryHelper.recordGauge("metric.gauge", 42.5);

        verify(telemetry).observeGauge("metric.gauge", 42.5);
    }

    @Test
    void testRecordRetryPromptedWithNullTags() {
        long startTime = System.currentTimeMillis() - 50;

        telemetryHelper.recordRetryPrompted("metric.test", startTime, null);

        verify(telemetry).incrementCounter(eq("metric.test.retry_prompted"), eq(1L));
        verify(telemetry).recordDuration(eq("metric.test.duration"), anyLong(), argThat(map ->
                "retry_prompted".equals(map.get("status"))
        ));
    }
}
