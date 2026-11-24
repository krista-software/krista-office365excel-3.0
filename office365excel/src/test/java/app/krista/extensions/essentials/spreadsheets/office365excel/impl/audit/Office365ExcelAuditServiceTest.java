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

package app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit;

import app.krista.extensions.util.KeyValueStore;
import app.krista.sdk.msauth.extension.audit.AuditData;
import org.junit.Before;
import org.junit.Test;

import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.Assert.*;

/**
 * Tests for Office365ExcelAuditService.
 * Covers the startRequest(invokerId, catalogName) overload
 * that records CATALOG_REQUEST_TRIGGERED.
 */
public class Office365ExcelAuditServiceTest {

    private MockKeyValueStore kvStore;
    private Office365ExcelAuditService service;

    private static final AtomicInteger COUNTER = new AtomicInteger(0);

    private String id() {
        return "excel-audit-test-" + COUNTER.incrementAndGet();
    }

    @Before
    public void setUp() {
        kvStore = new MockKeyValueStore();
        service = new Office365ExcelAuditService(kvStore);
    }

    // -- startRequest(invokerId, catalogName) ---------------------------------

    @Test
    public void testStartRequest_withInvokerAndCatalog_recordsTriggeredEvent() {
        String invokerId = id();
        service.startRequest(invokerId, "Read Range");

        AuditData data = service.getData(invokerId, null, null, null, null);
        assertNotNull(data);
        assertEquals(1, data.records.size());
        assertEquals("CATALOG_REQUEST_TRIGGERED", data.records.get(0).eventType);
        assertEquals("Read Range", data.records.get(0).catalogName);
        assertEquals("triggered", data.records.get(0).status);
    }

    @Test
    public void testStartRequest_withInvokerAndCatalog_thenSuccess_stampsDuration() throws InterruptedException {
        String invokerId = id();
        service.startRequest(invokerId, "Write Range");
        Thread.sleep(5);
        service.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", "Write Range", null, null, null);

        AuditData data = service.getData(invokerId, null, null, null, null);
        assertEquals(2, data.records.size());
        long totalDuration = data.records.stream().mapToLong(r -> r.durationMs).sum();
        assertTrue("Duration should be stamped on records", totalDuration > 0);
    }

    @Test
    public void testStartRequest_withInvokerAndCatalog_thenFailure_recordsBoth() {
        String invokerId = id();
        service.startRequest(invokerId, "Delete Row");
        service.record(invokerId, "CATALOG_REQUEST_FAILURE", "failure", "Delete Row",
                "Range not found", "IllegalArgumentException", null);

        AuditData data = service.getData(invokerId, null, null, null, null);
        assertEquals(2, data.records.size());
        assertEquals("CATALOG_REQUEST_TRIGGERED", data.records.get(0).eventType);
        assertEquals("CATALOG_REQUEST_FAILURE", data.records.get(1).eventType);
        assertEquals("Range not found", data.records.get(1).reason);
    }

    @Test
    public void testStartRequest_noArgs_doesNotRecordTriggered() {
        String invokerId = id();
        service.startRequest();
        // startRequest() alone should NOT record any event
        AuditData data = service.getData(invokerId, null, null, null, null);
        assertTrue("No events should be recorded by startRequest() alone",
                data.records == null || data.records.isEmpty());
    }

    @Test
    public void testRecord_directCall_works() {
        String invokerId = id();
        service.record(invokerId, "CATALOG_REQUEST_SUCCESS", "success", "Test", null, null, null);

        AuditData data = service.getData(invokerId, null, null, null, null);
        assertEquals(1, data.records.size());
        assertEquals("CATALOG_REQUEST_SUCCESS", data.records.get(0).eventType);
    }

    // -- MockKeyValueStore ----------------------------------------------------

    private static class MockKeyValueStore implements KeyValueStore {
        private final Map<String, Object> data = new ConcurrentHashMap<>();

        @Override public Object get(String key) { return data.get(key); }
        @Override
        @SuppressWarnings("unchecked")
        public <T> T get(String key, Class<T> type) {
            Object value = data.get(key);
            if (value == null) return null;
            if (type.isInstance(value)) return type.cast(value);
            return (T) value;
        }
        @Override public void remove(String key) { data.remove(key); }
        @Override public boolean contains(String key) { return data.containsKey(key); }
        @Override public void put(String key, Object value) { data.put(key, value); }
        @Override public boolean supports(Class<?> clazz) { return true; }
        @Override public Set<String> getKeys() { return new HashSet<>(data.keySet()); }
        @Override public Iterator<Map.Entry<String, Object>> iterator() { return data.entrySet().iterator(); }
    }
}
