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
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.audit.Office365ExcelAuditService;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.ksdk.context.AuthorizationContext;
import app.krista.ksdk.context.RequestContext;
import app.krista.model.field.NamedValuedField;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import javax.validation.constraints.NotNull;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

@Service
public class GraphServiceClientProviderFactory {

    private static final Logger LOGGER = LoggerFactory.getLogger(GraphServiceClientProviderFactory.class);

    private final RefreshTokenStore refreshTokenStore;
    private final OfficeExcelAttributeStore officeExcelAttributeStore;
    private final RequestContext requestContext;
    private final AuthorizationContext authorizationContext;
    private final Invoker invoker;
    private final GraphServiceClientProvider defaultClientProvider;
    private final Office365ExcelAuditService auditService;


    @Inject
    public GraphServiceClientProviderFactory(RefreshTokenStore refreshTokenStore,
                                             OfficeExcelAttributeStore officeExcelAttributeStore,
                                             RequestContext requestContext,
                                             AuthorizationContext authorizationContext,
                                             Invoker invoker,
                                             Office365ExcelAuditService auditService) {
        this.refreshTokenStore = refreshTokenStore;
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.invoker = invoker;
        this.auditService = auditService;
        this.defaultClientProvider = null;
    }

    public GraphServiceClientProviderFactory(RefreshTokenStore refreshTokenStore,
                                             OfficeExcelAttributeStore officeExcelAttributeStore,
                                             RequestContext requestContext,
                                             AuthorizationContext authorizationContext,
                                             Invoker invoker,
                                             GraphServiceClientProvider defaultClientProvider) {
        this.refreshTokenStore = refreshTokenStore;
        this.officeExcelAttributeStore = officeExcelAttributeStore;
        this.requestContext = requestContext;
        this.authorizationContext = authorizationContext;
        this.invoker = invoker;
        this.auditService = null;
        this.defaultClientProvider = defaultClientProvider;
    }

    public GraphServiceClientProvider create(OfficeExcelAttributes officeExcelAttributes) {
        String authContextId = officeExcelAttributeStore.save(officeExcelAttributes);
        return new GraphServiceClientProvider(refreshTokenStore, officeExcelAttributes, officeExcelAttributeStore, invoker,
                requestContext, authorizationContext, authContextId);
    }

    public GraphServiceClientProvider create(String authContextId) {
        final OfficeExcelAttributes attributes = officeExcelAttributeStore.load(authContextId);
        return buildProvider(attributes, authContextId);
    }

    public GraphServiceClientProvider create() {
        return create(requestContext);
    }

    public GraphServiceClientProvider create(RequestContext context) {
        // Rule 2: invokeAsUser check MUST happen FIRST — before any read of operator state.
        // If operator attrs are loaded first, a per-user request silently degrades to the
        // operator's identity when Setup credentials exist (data bleed).
        if (context.invokeAsUser()) {
            return createPerUserProvider(context);
        }

        // Operator-shared path: keyed by invokerId
        String key = context.getInvokerId();
        OfficeExcelAttributes attributes = officeExcelAttributeStore.load(key);
        if (attributes != null) {
            return buildProvider(attributes, null);
        }
        List<NamedValuedField> details = new ArrayList<>();
        details.add(new NamedValuedField(Constants.USER_ID, Constants.TEXT, key, new HashMap<>(), new HashMap<>()));
        throw new MustAuthorizeException("No credentials configured. Please set up credentials in the Authentication tab.", details);
    }

    /**
     * Per-user (MCP) auth path — completely isolated from operator state.
     * Looks up credentials by the user's wsContact accountId + suffix (Rule 1: hermetic keyspace).
     * If no per-user credentials exist, throws MustAuthorizeException to trigger consent (Rule 3).
     */
    private GraphServiceClientProvider createPerUserProvider(RequestContext context) {
        String wsContactId = authorizationContext.getAuthorizedAccount().getAccountId();
        String attrsKey = wsContactId + Constants.PER_USER_ATTRS_SUFFIX;
        LOGGER.info("Per-user auth: wsContactId={}, attrsKey={}, invokerId={}", wsContactId, attrsKey, context.getInvokerId());
        OfficeExcelAttributes attrs = officeExcelAttributeStore.load(attrsKey);
        if (attrs != null) {
            LOGGER.info("Per-user auth: found saved attrs for {} — email={}, clientId={}, authType={}",
                    wsContactId, attrs.getEmail(), attrs.getClientId(), attrs.getAuthType());
            return buildProvider(attrs, null);
        }

        // No per-user credentials — trigger OAuth consent flow
        try {
            List<NamedValuedField> details = new ArrayList<>();
            details.add(new NamedValuedField(Constants.CLIENT_USER_ID, Constants.TEXT, wsContactId, new HashMap<>(), new HashMap<>()));
            details.add(new NamedValuedField(Constants.USER_ID, Constants.TEXT, context.getInvokerId(), new HashMap<>(), new HashMap<>()));
            throw new MustAuthorizeException("No credentials configured. Please authenticate via the Authentication tab.", details);
        } catch (MustAuthorizeException e) {
            throw e;
        } catch (Exception e) {
            // Rule 3: never silently fall through to operator path — always propagate as MAE
            LOGGER.error("Per-user auth setup failed: {}", e.getMessage(), e);
            throw new MustAuthorizeException("Per-user authentication setup failed. Please try again.", new ArrayList<>());
        }
    }

    public String createAttributes(OfficeExcelAttributes attributes) {
        return officeExcelAttributeStore.save(attributes);
    }

    @NotNull
    private GraphServiceClientProvider buildProvider(OfficeExcelAttributes attributes,
                                                     String authContextId) {
        return new GraphServiceClientProvider(refreshTokenStore, attributes, officeExcelAttributeStore, invoker, requestContext, authorizationContext,
                authContextId);
    }



}
