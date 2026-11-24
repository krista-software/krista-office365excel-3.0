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

package app.krista.extensions.essentials.spreadsheets.office365excel.api;

import app.krista.extension.executor.Invoker;
import app.krista.extension.request.RoutingInfo;
import app.krista.extension.request.protos.http.HttpProtocol;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import javax.ws.rs.*;
import javax.ws.rs.core.Response;
import java.net.URI;

import static com.github.scribejava.core.model.OAuthConstants.CODE;
import static com.github.scribejava.core.model.OAuthConstants.STATE;

/**
 * REST resource for handling generic OAuth callbacks for public authentication.
 * This resource handles the generic /extension/api/rest/v3/oauth/callback endpoint
 * and forwards requests to the appropriate extension-specific callback endpoint.
 */
@Path("/")
public final class RestCallbackResource {
    private static final Logger LOGGER = LoggerFactory.getLogger(RestCallbackResource.class);
    
    private final String baseRoutingUrl;

    @Inject
    public RestCallbackResource(Invoker invoker) {
        this.baseRoutingUrl = invoker.getRoutingInfo().getRoutingURL(HttpProtocol.PROTOCOL_NAME, RoutingInfo.Type.APPLIANCE);
    }

    /**
     * Handles OAuth callback requests for v3 authentication flow.
     * This endpoint receives callbacks from Microsoft OAuth and forwards them
     * to the extension-specific callback endpoint based on the state parameter.
     *
     * @param code  OAuth authorization code received from Microsoft
     * @param state State parameter containing user and context information
     * @return Redirect response to the extension-specific callback endpoint
     */
    @GET
    @Path("v3/oauth/callback")
    public Response handleGenericCallback(@QueryParam(CODE) String code, @QueryParam(STATE) String state) {
        try {
            LOGGER.info("Received generic OAuth callback with code: {} and state: {}", 
                       code != null ? "present" : "null", state != null ? "present" : "null");
            
            // Extract forward path from state parameter
            String forwardPath = extractForwardPath(state);
            if (forwardPath == null) {
                LOGGER.error("No forward path found in state parameter: {}", state);
                return Response.status(Response.Status.BAD_REQUEST)
                              .entity("Invalid state parameter - no forward path found")
                              .build();
            }
            
            // Construct the forward URL
            String forwardUrl = baseRoutingUrl + forwardPath + "?code=" + code + "&state=" + state;
            
            LOGGER.info("Forwarding OAuth callback to: {}", forwardUrl);
            
            // Redirect to the extension-specific callback endpoint
            return Response.status(Response.Status.FOUND)
                          .location(URI.create(forwardUrl))
                          .build();
                          
        } catch (Exception e) {
            LOGGER.error("Error handling generic OAuth callback", e);
            return Response.status(Response.Status.INTERNAL_SERVER_ERROR)
                          .entity("Error processing OAuth callback: " + e.getMessage())
                          .build();
        }
    }
    
    /**
     * Extracts the forward path from the state parameter.
     * The state parameter format is: userId#authContextId#forwardPath
     *
     * @param state The state parameter from OAuth callback
     * @return The forward path or null if not found
     */
    private String extractForwardPath(String state) {
        if (state == null || state.isEmpty()) {
            LOGGER.error("State parameter is null or empty");
            return null;
        }

        String[] parts = state.split("#");
        LOGGER.info("State parameter parts: {}", java.util.Arrays.toString(parts));

        if (parts.length >= 3) {
            String forwardPath = parts[2];
            LOGGER.info("Extracted forward path: {}", forwardPath);
            return forwardPath; // Forward path is the third part
        }

        LOGGER.error("State parameter does not contain forward path. Expected format: userId#authContextId#forwardPath, got: {}", state);
        return null;
    }
}
