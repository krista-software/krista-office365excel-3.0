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

package app.krista.extensions.essentials.spreadsheets.office365excel;

import app.krista.extension.authorization.MustAuthorizeException;
import app.krista.extension.authorization.RequestAuthenticator;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import app.krista.sdk.msauth.MicrosoftAuthSdk;
import app.krista.sdk.msauth.extension.base.AbstractRequestAuthenticator;
import app.krista.sdk.msauth.extension.base.MsAuthAttributes;
import app.krista.sdk.msauth.token.ITokenStore;
import app.krista.sdk.msauth.token.InMemoryTokenStore;
import app.krista.model.field.NamedValuedField;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.Collections;
import java.util.List;

import static app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants.PUBLIC;

/**
 * Office 365 Excel-specific request authenticator.
 *
 * <p>Extends {@link AbstractRequestAuthenticator} which uses {@code MicrosoftAuthSdk}
 * directly for OAuth URL generation — no ScribeJava or OAuthService needed.</p>
 */
@Service
public class OfficeExcelRequestAuthenticator extends AbstractRequestAuthenticator {

    private static final Logger LOGGER = LoggerFactory.getLogger(OfficeExcelRequestAuthenticator.class);

    private final OfficeExcelAttributeStore attributeStore;

    @Inject
    public OfficeExcelRequestAuthenticator(OfficeExcelAttributeStore attributeStore, String invokerId,
                                           String routingUrl) {
        super(invokerId, routingUrl, "/excel365extn");
        this.attributeStore = attributeStore;
    }

    @Override
    protected MsAuthAttributes loadAttributes(String key) {
        return attributeStore.load(key);
    }

    @Override
    protected ITokenStore getTokenStore() {
        // Authenticator only builds OAuth URLs — doesn't need real token persistence.
        // The actual token store is used by GraphServiceClientProviderFactory.
        return new InMemoryTokenStore();
    }

    @Override
    protected MsAuthAttributes createDefaultAttributes(String routingUrl) {
        return new OfficeExcelAttributes(null, null, null, null, PUBLIC, routingUrl);
    }

    @Override
    public RequestAuthenticator.AuthorizationResponse getMustAuthorizeResponse(MustAuthorizeException cause) {
        List<NamedValuedField> details = cause.getDetails();
        if (details == null || details.isEmpty()) {
            throw new IllegalStateException("MustAuthorizeException has no details");
        }
        String fieldName = details.get(0).getName();
        String userId    = (String) details.get(0).getValue();
        LOGGER.info("getMustAuthorizeResponse: fieldName={}, userId={}, detailCount={}",
                fieldName, userId, details.size());

        if (userId != null && userId.startsWith(Constants.WS_CONTACT)) {
            LOGGER.info("getMustAuthorizeResponse: per-user (wsContact) — building OAuth redirect");
            MsAuthAttributes attrs = createDefaultAttributes(getRoutingUrl());
            MicrosoftAuthSdk sdk = attrs.toSdk(getTokenStore());
            String state = userId + Constants.HASH + getInvokerId();
            if (Constants.PUBLIC.equals(attrs.getAuthType())) {
                state += Constants.HASH + attrs.getForwardPath();
            }
            LOGGER.info("getMustAuthorizeResponse: state built — wsContactId#invokerId={}#{}", userId, getInvokerId());
            String url = sdk.buildAuthorizationUrl(state);
            LOGGER.info("getMustAuthorizeResponse: OAuth URL generated");
            return new RequestAuthenticator.AuthorizationResponse(url, Collections.emptyList());
        }

        LOGGER.info("getMustAuthorizeResponse: non-wsContact userId={} — delegating to super", userId);
        return super.getMustAuthorizeResponse(cause);
    }
}
