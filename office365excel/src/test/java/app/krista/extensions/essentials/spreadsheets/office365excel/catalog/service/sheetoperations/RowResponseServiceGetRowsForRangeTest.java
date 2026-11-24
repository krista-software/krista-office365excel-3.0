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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.sheetoperations;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.ExtensionResponse;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.SheetFormattingService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.base.FreeForm;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.ArgumentCaptor;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * 100% branch coverage for:
 * <ul>
 *   <li>{@link RowResponseService#getGetRowsForRangeExtensionResponse(String, String, String, Boolean)}
 *       — 4-arg overload that delegates to the 5-arg version with driveId=null</li>
 *   <li>{@link RowResponseService#getGetRowsForRangeExtensionResponse(String, String, String, String, Boolean)}
 *       — 5-arg implementation that builds and executes an {@code AreaOperationHandler}</li>
 * </ul>
 *
 * <p>Branches covered:
 * <ul>
 *   <li>4-arg → driveId=null forwarded to worksheetService.getRows and StateMapperUtil</li>
 *   <li>telemetryHelper.incrementCount("office365Excel.getRowsForRange") called on every invocation</li>
 *   <li>allowRetry=null → validation skipped → operation executed directly</li>
 *   <li>allowRetry=false → validation skipped → operation executed directly</li>
 *   <li>allowRetry=true + validation passes → operation executed</li>
 *   <li>allowRetry=true + validation fails → CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET used,
 *       state stored, operation never called</li>
 *   <li>allowRetry=true + validation fails → state map contains driveId, fileId, sheetName, range, stateId</li>
 *   <li>allowRetry=true + validation fails + driveId=null → null driveId propagated in state map</li>
 *   <li>allowRetry=true + validation throws IllegalStateException → error response, operation never called</li>
 *   <li>worksheetService.getRows throws MustAuthorizeException → auth response returned</li>
 *   <li>worksheetService.getRows throws RuntimeException → state removed, error response returned</li>
 *   <li>non-null driveId (5-arg) forwarded verbatim to worksheetService.getRows</li>
 * </ul>
 */
@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
@DisplayName("RowResponseService – getGetRowsForRangeExtensionResponse – 100% branch coverage")
class RowResponseServiceGetRowsForRangeTest {

    // ─── Test constants ──────────────────────────────────────────────────────
    private static final String FILE_ID    = "file-123";
    private static final String DRIVE_ID   = "drive-456";
    private static final String SHEET_NAME = "Sheet1";
    private static final String RANGE      = "A1:D10";

    // ─── Mocks ───────────────────────────────────────────────────────────────
    @Mock private WorksheetService        worksheetService;
    @Mock private ValidationOrchestrator  validationOrchestrator;
    @Mock private ErrorHandlingStateManager internalStateManager;
    @Mock private ExtensionResponseGenerator responseGenerator;
    @Mock private RequestContext          requestContext;
    @Mock private TelemetryHelper         telemetryHelper;
    @Mock private RestClientProvider      restClientProvider;
    @Mock private StringUtil              stringUtil;
    @Mock private SheetFormattingService  sheetFormattingService;

    private RowResponseService service;

    @BeforeEach
    void setUp() {
        service = new RowResponseService(
                worksheetService, validationOrchestrator, internalStateManager,
                responseGenerator, requestContext, telemetryHelper,
                restClientProvider, stringUtil, sheetFormattingService);
    }

    // ─── Helpers ─────────────────────────────────────────────────────────────

    /** Two-element row list returned by worksheetService on the happy path. */
    private List<FreeForm> fakeRows() {
        return List.of(mock(FreeForm.class), mock(FreeForm.class));
    }

    /** Single validation error — enough to trigger the confirmation-response branch. */
    private List<ValidationOrchestrator.ValidationResult> oneValidationError() {
        return List.of(new ValidationOrchestrator.ValidationResult(
                "confirmMsg", "File Id", "fetchMsg", "File not found", "TEXT"));
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 4-arg overload — delegates to 5-arg with driveId=null
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("4-arg overload — delegates to 5-arg with driveId=null")
    class FourArgDelegation {

        @Test
        @DisplayName("null driveId forwarded to worksheetService.getRows")
        void fourArgOverload_forwardsNullDriveId() {
            when(worksheetService.getRows(null, FILE_ID, SHEET_NAME, RANGE)).thenReturn(fakeRows());

            ExtensionResponse result = service.getGetRowsForRangeExtensionResponse(
                    FILE_ID, SHEET_NAME, RANGE, null);

            assertNotNull(result);
            verify(worksheetService).getRows(null, FILE_ID, SHEET_NAME, RANGE);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Telemetry count
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Telemetry — incrementCount on every invocation")
    class TelemetryCount {

        @Test
        @DisplayName("4-arg overload increments \"office365Excel.getRowsForRange\"")
        void fourArg_incrementsCount() {
            when(worksheetService.getRows(null, FILE_ID, SHEET_NAME, RANGE)).thenReturn(fakeRows());

            service.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, null);

            verify(telemetryHelper).incrementCount("office365Excel.getRowsForRange");
        }

        @Test
        @DisplayName("5-arg overload increments \"office365Excel.getRowsForRange\"")
        void fiveArg_incrementsCount() {
            when(worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE)).thenReturn(fakeRows());

            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

            verify(telemetryHelper).incrementCount("office365Excel.getRowsForRange");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // allowRetry branches — skip vs run validation
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("allowRetry branches — validation skipped vs executed")
    class AllowRetryBranches {

        @BeforeEach
        void stubWorksheetService() {
            when(worksheetService.getRows(eq(DRIVE_ID), eq(FILE_ID), eq(SHEET_NAME), eq(RANGE)))
                    .thenReturn(fakeRows());
        }

        @Test
        @DisplayName("allowRetry=null → validation NOT called, worksheetService.getRows IS called")
        void allowRetryNull_skipsValidation_callsGetRows() {
            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

            verify(validationOrchestrator, never()).validate(anyMap());
            verify(worksheetService).getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);
        }

        @Test
        @DisplayName("allowRetry=false → validation NOT called, worksheetService.getRows IS called")
        void allowRetryFalse_skipsValidation_callsGetRows() {
            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, false);

            verify(validationOrchestrator, never()).validate(anyMap());
            verify(worksheetService).getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);
        }

        @Test
        @DisplayName("allowRetry=true + validation passes → validation IS called, worksheetService.getRows IS called")
        void allowRetryTrue_validationPasses_callsGetRows() {
            when(validationOrchestrator.validate(anyMap())).thenReturn(List.of());

            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

            verify(validationOrchestrator).validate(anyMap());
            verify(worksheetService).getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Validation failure branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Validation failure branches (allowRetry=true + validation fails)")
    class ValidationFailureBranches {

        @BeforeEach
        void stubValidationFailure() {
            when(validationOrchestrator.validate(anyMap())).thenReturn(oneValidationError());
            when(responseGenerator.generateConfirmationResponse(any(), anyList(), anyString(), anyMap()))
                    .thenReturn(mock(ExtensionResponse.class));
        }

        @Test
        @DisplayName("confirmation action is CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET")
        void confirmationAction_isCorrectConstant() {
            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

            verify(responseGenerator).generateConfirmationResponse(
                    any(),
                    anyList(),
                    eq(SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET),
                    anyMap());
        }

        @Test
        @DisplayName("state is stored in internalStateManager and worksheetService never called")
        void stateStored_operationSkipped() {
            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

            verify(internalStateManager).put(anyString(), anyMap());
            verify(worksheetService, never()).getRows(any(), any(), any(), any());
        }

        @Test
        @DisplayName("state map passed to responseGenerator contains driveId, fileId, sheetName, range, stateId")
        void stateMap_containsAllExpectedKeys_withNonNullDriveId() {
            @SuppressWarnings("unchecked")
            ArgumentCaptor<Map<String, Object>> captor = ArgumentCaptor.forClass(Map.class);
            when(responseGenerator.generateConfirmationResponse(any(), anyList(), anyString(), captor.capture()))
                    .thenReturn(mock(ExtensionResponse.class));

            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

            Map<String, Object> stateMap = captor.getValue();
            assertEquals(DRIVE_ID,    stateMap.get("Drive Id"));
            assertEquals(FILE_ID,     stateMap.get("File Id"));
            assertEquals(SHEET_NAME,  stateMap.get("Sheet Name"));
            assertEquals(RANGE,       stateMap.get("Range"));
            assertNotNull(stateMap.get("State Id")); // random UUID — just verify presence
        }

        @Test
        @DisplayName("4-arg overload (driveId=null) → null driveId propagated in state map")
        void fourArgOverload_nullDriveId_propagatedInStateMap() {
            @SuppressWarnings("unchecked")
            ArgumentCaptor<Map<String, Object>> captor = ArgumentCaptor.forClass(Map.class);
            when(responseGenerator.generateConfirmationResponse(any(), anyList(), anyString(), captor.capture()))
                    .thenReturn(mock(ExtensionResponse.class));

            // 4-arg overload → calls 5-arg with driveId=null
            service.getGetRowsForRangeExtensionResponse(FILE_ID, SHEET_NAME, RANGE, true);

            Map<String, Object> stateMap = captor.getValue();
            assertNull(stateMap.get("Drive Id"),
                    "driveId=null from 4-arg overload must not appear in the state map");
        }

        @Test
        @DisplayName("confirmation response returned directly without calling worksheetService")
        void confirmationResponse_returnedDirectly() {
            ExtensionResponse expected = mock(ExtensionResponse.class);
            when(responseGenerator.generateConfirmationResponse(any(), anyList(),
                    eq(SubCatalogConstants.CONFIRM_REENTER_GET_ROWS_FOR_RANGE_FOR_SHEET), anyMap()))
                    .thenReturn(expected);

            ExtensionResponse result =
                    service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

            assertSame(expected, result);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Validation throws IllegalStateException
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Validation throws IllegalStateException (access/infrastructure error)")
    class ValidationIllegalStateException {

        @Test
        @DisplayName("error response returned, worksheetService never called")
        void illegalStateException_returnsErrorResponse_operationSkipped() {
            when(validationOrchestrator.validate(anyMap()))
                    .thenThrow(new IllegalStateException("access denied"));

            ExtensionResponse result =
                    service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, true);

            assertNotNull(result);
            verify(worksheetService, never()).getRows(any(), any(), any(), any());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Operation exception branches
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("Operation exception branches")
    class OperationExceptionBranches {

        @Test
        @DisplayName("worksheetService.getRows throws MustAuthorizeException → non-null response returned")
        void mustAuthorizeException_returnsNonNullResponse() {
            MustAuthorizeException ex = mock(MustAuthorizeException.class);
            when(worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE)).thenThrow(ex);
            when(requestContext.invokeAsUser()).thenReturn(false);

            ExtensionResponse result =
                    service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

            assertNotNull(result);
        }

        @Test
        @DisplayName("worksheetService.getRows throws RuntimeException → state removed, non-null error response")
        void runtimeException_removesState_returnsErrorResponse() {
            when(worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE))
                    .thenThrow(new RuntimeException("Graph API error"));

            ExtensionResponse result =
                    service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

            assertNotNull(result);
            verify(internalStateManager).remove(anyString());
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // driveId propagation to worksheetService
    // ═══════════════════════════════════════════════════════════════════════

    @Nested
    @DisplayName("driveId propagation to worksheetService.getRows")
    class DriveIdPropagation {

        @Test
        @DisplayName("non-null driveId forwarded verbatim to worksheetService.getRows")
        void nonNullDriveId_forwardedVerbatim() {
            when(worksheetService.getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE)).thenReturn(fakeRows());

            service.getGetRowsForRangeExtensionResponse(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE, null);

            verify(worksheetService).getRows(DRIVE_ID, FILE_ID, SHEET_NAME, RANGE);
        }

        @Test
        @DisplayName("null driveId (5-arg) forwarded as null to worksheetService.getRows")
        void nullDriveId_fiveArg_forwardedAsNull() {
            when(worksheetService.getRows(null, FILE_ID, SHEET_NAME, RANGE)).thenReturn(fakeRows());

            service.getGetRowsForRangeExtensionResponse((String) null, FILE_ID, SHEET_NAME, RANGE, null);

            verify(worksheetService).getRows(isNull(), eq(FILE_ID), eq(SHEET_NAME), eq(RANGE));
        }
    }
}
