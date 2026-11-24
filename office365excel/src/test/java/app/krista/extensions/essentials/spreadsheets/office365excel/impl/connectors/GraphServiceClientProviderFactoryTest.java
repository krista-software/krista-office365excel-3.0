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

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.accounts.Account;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class GraphServiceClientProviderFactoryTest {

    @Mock private RefreshTokenStore refreshTokenStore;
    @Mock private OfficeExcelAttributeStore officeExcelAttributeStore;
    @Mock private RequestContext requestContext;
    @Mock private AuthorizationContext authorizationContext;
    @Mock private Invoker invoker;

    private GraphServiceClientProviderFactory factory;

    private static final String INVOKER_ID = "INV_testInvoker";

    @BeforeEach
    void setUp() {
        when(requestContext.getInvokerId()).thenReturn(INVOKER_ID);
        factory = new GraphServiceClientProviderFactory(
                refreshTokenStore, officeExcelAttributeStore, requestContext,
                authorizationContext, invoker, (GraphServiceClientProvider) null);
    }

    // ── invokeAsUser=true, attrs found → returns provider without throwing ─────

    @Test
    void create_invokeAsUser_attrsFound_returnsProvider_withoutThrowing() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(requestContext.invokeAsUser()).thenReturn(true);

        // Per-user path: factory checks invokeAsUser FIRST, then loads by wsContactId + "_excel_attrs"
        String wsContactId = "wsContactId_abc123";
        Account account = mock(Account.class);
        when(account.getAccountId()).thenReturn(wsContactId);
        when(authorizationContext.getAuthorizedAccount()).thenReturn(account);
        when(officeExcelAttributeStore.load(wsContactId + "_excel_attrs")).thenReturn(attrs);

        assertDoesNotThrow(() -> factory.create(requestContext));
    }

    // ── invokeAsUser=true, no attrs → throws MAE with wsContact detail ─────────

    @Test
    void create_invokeAsUser_noAttrs_throwsMustAuthorizeException_withWsContact() {
        when(requestContext.invokeAsUser()).thenReturn(true);

        Account account = mock(Account.class);
        when(account.getAccountId()).thenReturn("wsContactId_abc123");
        when(authorizationContext.getAuthorizedAccount()).thenReturn(account);
        when(officeExcelAttributeStore.load("wsContactId_abc123" + "_excel_attrs")).thenReturn(null);

        MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                () -> factory.create(requestContext));

        assertNotNull(ex.getDetails());
        assertFalse(ex.getDetails().isEmpty());
        String firstFieldName = ex.getDetails().get(0).getName();
        String firstValue = (String) ex.getDetails().get(0).getValue();
        assertEquals(Constants.CLIENT_USER_ID, firstFieldName);
        assertTrue(firstValue.startsWith("wsContact"));
    }

    // ── invokeAsUser=false, attrs found → returns provider ─────────────────────

    @Test
    void create_notInvokeAsUser_attrsFound_returnsProvider() {
        OfficeExcelAttributes attrs = mock(OfficeExcelAttributes.class);
        when(requestContext.invokeAsUser()).thenReturn(false);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(attrs);
        when(officeExcelAttributeStore.save(any(OfficeExcelAttributes.class))).thenReturn("ctx-2");

        assertDoesNotThrow(() -> factory.create(requestContext));
    }

    // ── invokeAsUser=false, no attrs → throws MAE with userId detail ────────────

    @Test
    void create_notInvokeAsUser_noAttrs_throwsMustAuthorizeException() {
        when(requestContext.invokeAsUser()).thenReturn(false);
        when(officeExcelAttributeStore.load(INVOKER_ID)).thenReturn(null);

        MustAuthorizeException ex = assertThrows(MustAuthorizeException.class,
                () -> factory.create(requestContext));

        assertNotNull(ex.getDetails());
        assertFalse(ex.getDetails().isEmpty());
        assertEquals(Constants.USER_ID, ex.getDetails().get(0).getName());
        assertEquals(INVOKER_ID, ex.getDetails().get(0).getValue());
    }
}
