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

package app.krista.extensions.essentials.spreadsheets.office365excel.mcp;

import app.krista.extension.executor.ExtensionResponse;
import app.krista.extension.executor.Invoker;
import app.krista.extension.executor.RemediationAction;
import app.krista.extension.executor.RemediationActions;
import app.krista.extension.executor.impl.InformAPersonAction;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.model.base.FreeForm;
import app.krista.sdk.msauth.extension.mcp.AbstractMcpQueryService;
import app.krista.sdk.msauth.extension.mcp.AnthropicApiClient;
import app.krista.sdk.msauth.extension.mcp.DefaultMsAuthBridge;
import app.krista.sdk.msauth.extension.mcp.GeminiApiClient;
import app.krista.sdk.msauth.extension.mcp.McpModelConfig;
import app.krista.sdk.msauth.extension.mcp.McpProductConfig;
import app.krista.sdk.msauth.extension.mcp.McpSearchContext;
import app.krista.sdk.msauth.extension.mcp.McpSessionKeys;
import app.krista.sdk.msauth.extension.mcp.OpenAiApiClient;
import app.krista.sdk.msauth.token.ITokenStore;
import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import javax.ws.rs.client.Client;
import javax.ws.rs.client.ClientBuilder;
import javax.ws.rs.client.Entity;
import javax.ws.rs.core.MediaType;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Office 365 Excel-specific MCP query service — extends the SDK's
 * {@link AbstractMcpQueryService} with dynamic AI-driven Graph API routing.
 *
 * <p>The {@link #directQuery(String)} entry point asks the AI to plan any
 * Microsoft Graph API call ({@code method}, {@code path}, {@code body}) and
 * executes it generically — no hardcoded intents.</p>
 *
 * <p>Used by {@link ExcelMcpSetupResource#interceptNlQuery(String)} for the
 * AI Assistant tab.</p>
 */
@Service
public class ExcelMcpQueryService extends AbstractMcpQueryService {

    private static final Logger LOG = LoggerFactory.getLogger(ExcelMcpQueryService.class);
    private static final Gson GSON = new Gson();

    private final ITokenStore tokenStore;
    private final Invoker invoker;
    private final OfficeExcelAttributeStore attributeStore;

    @Inject
    public ExcelMcpQueryService(RefreshTokenStore refreshTokenStore,
                                Invoker invoker,
                                OfficeExcelAttributeStore attributeStore) {
        super(adaptTokenStore(refreshTokenStore),
              invoker,
              new DefaultMsAuthBridge(
                      attributeStore.load(invoker.getInvokerId()),
                      adaptTokenStore(refreshTokenStore)));
        this.tokenStore   = adaptTokenStore(refreshTokenStore);
        this.invoker      = invoker;
        this.attributeStore = attributeStore;
    }

    /**
     * Reloads credentials from the attribute store on every call so that
     * changes in the Authentication tab take effect without a restart.
     */
    private DefaultMsAuthBridge currentAuthBridge() {
        OfficeExcelAttributes attrs = attributeStore.load(invoker.getInvokerId());
        return new DefaultMsAuthBridge(attrs, tokenStore);
    }

    /** Returns current attributes, reloaded on every call. */
    private OfficeExcelAttributes currentAttributes() {
        return attributeStore.load(invoker.getInvokerId());
    }

    @Override
    protected McpProductConfig getProductConfig() {
        return new ExcelMcpProductConfig();
    }

    @Override
    protected String getJaxrsId() {
        return "excel";
    }

    /**
     * Single entry point for both catalog requests and the NQL tab.
     *
     * <p>Asks the AI to plan the exact Microsoft Graph API call ({@code method},
     * {@code path}, {@code body}) needed to answer the user's question, then
     * executes it generically via {@link #executeGraphCall}. No hardcoded
     * intents — any valid Graph endpoint works without code changes.</p>
     *
     * <p>Falls back to the parent KQL search pipeline if the AI returns an
     * unparseable response.</p>
     */
    @Override
    public ExtensionResponse query(String question) {
        long startTime = System.currentTimeMillis();
        String id = invoker.getInvokerId();
        LOG.info("[MCP-XL] query START — invokerId: {}, question: '{}'", id, question);

        try {
            // Resolve AI provider + API key
            String provider = tokenStore.get(McpSessionKeys.aiProvider(id));
            String aiKey    = blank(provider) ? null : tokenStore.get(McpSessionKeys.aiKey(id, provider));

            if (blank(aiKey)) {
                aiKey = tokenStore.get(McpSessionKeys.claudeApiKey(id));
                if (!blank(aiKey)) provider = "anthropic";
            }

            if (blank(aiKey)) {
                LOG.warn("[MCP-XL] query — No AI key configured");
                return failure("AI API key is not configured. "
                        + "Please set it in the AI Assistant tab \u2192 AI Settings.");
            }

            // Reload auth on every call so credential changes take effect immediately
            DefaultMsAuthBridge authBridge = currentAuthBridge();
            if (!authBridge.isAuthenticated()) {
                return failure("Not connected. Configure credentials in the Authentication tab.");
            }

            // Ask AI to plan the Graph API call
            ExcelMcpProductConfig config = (ExcelMcpProductConfig) getProductConfig();
            String planPrompt = config.graphApiPlanningPrompt(question);
            LOG.info("[MCP-XL] query — planning Graph API call via AI...");
            String raw = callAi(aiKey, provider, planPrompt);
            LOG.info("[MCP-XL] query — AI plan: '{}'", raw);

            if (blank(raw)) {
                LOG.warn("[MCP-XL] query — AI returned blank response");
                return failure("Could not determine Graph API call. Please try rephrasing your question.");
            }

            if (raw.contains("{")) {
                raw = raw.substring(raw.indexOf('{'), raw.lastIndexOf('}') + 1);
            }

            JsonObject plan   = GSON.fromJson(raw, JsonObject.class);
            String method     = getJsonString(plan, "method");
            String path       = getJsonString(plan, "path");
            String body       = plan != null && plan.has("body") && !plan.get("body").isJsonNull()
                    ? (plan.get("body").isJsonPrimitive()
                            ? plan.get("body").getAsString()
                            : plan.get("body").toString())
                    : null;

            if (blank(method) || blank(path)) {
                LOG.warn("[MCP-XL] query — AI returned invalid plan: {}", raw);
                return failure("Could not parse a valid Graph API call from the AI response. Please rephrase your question.");
            }

            LOG.info("[MCP-XL] query — executing {} {}", method, path);

            McpSearchContext ctx   = authBridge.buildSearchContext(config);
            String authHeader      = ctx.authHeader();
            String resolvedPathStr = resolvedPath(path, authBridge);

            // Resolve {itemId} placeholder by searching for the file by name
            if (resolvedPathStr.contains("{itemId}")) {
                String pathWithId = resolveItemId(resolvedPathStr, question, authHeader);
                if (pathWithId == null) {
                    return failure("Could not find the Excel file mentioned in your question. "
                            + "Please check the file name and try again.");
                }
                resolvedPathStr = pathWithId;
            }

            String fullUrl = "https://graph.microsoft.com/v1.0" + resolvedPathStr;
            ExtensionResponse result = executeGraphCall(method, fullUrl, body, authHeader);
            // Attach the AI-generated Graph plan string so the UI can display it
            if (result.getResult() == ExtensionResponse.Result.SUCCESS
                    && result.getResponseValue() instanceof Map) {
                @SuppressWarnings("unchecked")
                Map<String, Object> rv = (Map<String, Object>) result.getResponseValue();
                rv.put("generatedPlan", method + " " + path);
            }
            LOG.info("[MCP-XL] query END — elapsed: {}ms",
                    System.currentTimeMillis() - startTime);
            return result;

        } catch (Exception e) {
            LOG.error("[MCP-XL] query FAILED — elapsed: {}ms",
                    System.currentTimeMillis() - startTime, e);
            return failure("Query failed: " + e.getMessage());
        }
    }

    /**
     * Replaces {@code /me/} with {@code /users/{email}/} when private (service-to-service)
     * auth is configured, so AI-generated paths work for both auth modes.
     * Uses fresh attributes so credential changes are reflected immediately.
     */
    private String resolvedPath(String path, DefaultMsAuthBridge authBridge) {
        OfficeExcelAttributes attrs = currentAttributes();
        if (attrs != null && "PRIVATE".equals(attrs.getAuthType())) {
            String email = attrs.getEmail();
            if (!blank(email)) {
                return path.replaceFirst("^/me(/|$)", "/users/" + email + "$1");
            }
        }
        return path;
    }

    /**
     * Executes any Microsoft Graph API call and converts the response into
     * a list of FreeForm rows for the NQL tab.
     *
     * <p>Handles GET and POST. HTTP 202/204 (no body) treated as success.
     * HTTP 200/201 responses are parsed by {@link #parseGenericResponse}.</p>
     */
    private ExtensionResponse executeGraphCall(String method, String fullUrl,
                                               String body, String authHeader) {
        Client client = ClientBuilder.newClient();
        try {
            javax.ws.rs.core.Response httpResp;
            if ("POST".equalsIgnoreCase(method)) {
                httpResp = client.target(fullUrl)
                        .request(MediaType.APPLICATION_JSON)
                        .header("Authorization", authHeader)
                        .post(Entity.entity(body != null ? body : "{}",
                                MediaType.APPLICATION_JSON));
            } else if ("PATCH".equalsIgnoreCase(method)) {
                httpResp = client.target(fullUrl)
                        .request(MediaType.APPLICATION_JSON)
                        .header("Authorization", authHeader)
                        .method("PATCH", Entity.entity(body != null ? body : "{}",
                                MediaType.APPLICATION_JSON));
            } else if ("DELETE".equalsIgnoreCase(method)) {
                httpResp = client.target(fullUrl)
                        .request(MediaType.APPLICATION_JSON)
                        .header("Authorization", authHeader)
                        .delete();
            } else {
                httpResp = client.target(fullUrl)
                        .request(MediaType.APPLICATION_JSON)
                        .header("Authorization", authHeader)
                        .get();
            }

            int status = httpResp.getStatus();
            LOG.info("[MCP-XL] executeGraphCall — {} {} → HTTP {}", method, fullUrl, status);

            if (status == 202 || status == 204) {
                FreeForm ff = new FreeForm();
                ff.put("Answer", " Text ", "Action completed successfully.");
                Map<String, Object> actionMap = new HashMap<>();
                actionMap.put("Response", List.of(ff));
                return success(actionMap);
            }

            if (status == 200 || status == 201) {
                String responseBody = httpResp.readEntity(String.class);

                // PATCH returns the updated range object — show a friendly success message
                if ("PATCH".equalsIgnoreCase(method)) {
                    String updatedAddress = "";
                    try {
                        JsonObject rangeObj = GSON.fromJson(responseBody, JsonObject.class);
                        updatedAddress = getStr(rangeObj, "address");
                    } catch (Exception ignored) {}
                    FreeForm ff = new FreeForm();
                    ff.put("Result", " Text ", "Cell updated successfully."
                            + (updatedAddress.isEmpty() ? "" : " (" + updatedAddress + ")"));
                    Map<String, Object> patchMap = new HashMap<>();
                    patchMap.put("Response", List.of(ff));
                    return success(patchMap);
                }

                List<FreeForm> rows = parseGenericResponse(responseBody);
                LOG.info("[MCP-XL] executeGraphCall — parsed {} row(s)", rows.size());
                Map<String, Object> responseMap = new HashMap<>();
                responseMap.put("Response", rows);
                return success(responseMap);
            }

            String errorBody = httpResp.readEntity(String.class);
            LOG.error("[MCP-XL] executeGraphCall FAILED — HTTP {}: {}", status, errorBody);
            return failure("Graph API call failed (HTTP " + status + "): " + errorBody);

        } catch (Exception e) {
            LOG.error("[MCP-XL] executeGraphCall FAILED — {}", e.getMessage(), e);
            return failure("Graph API call failed: " + e.getMessage());
        } finally {
            client.close();
        }
    }

    /**
     * Resolves the {@code {itemId}} placeholder in a Graph API path by searching
     * for the file by name extracted from the user's question.
     *
     * @return the path with {@code {itemId}} replaced by the real drive item ID,
     *         or {@code null} if the file could not be found
     */
    private String resolveItemId(String path, String question, String authHeader) {
        String fileName = extractFileName(question);
        if (fileName == null || fileName.isEmpty()) {
            LOG.warn("[MCP-XL] resolveItemId — could not extract filename from question: '{}'", question);
            return null;
        }

        LOG.info("[MCP-XL] resolveItemId — searching for file: '{}'", fileName);
        String searchUrl = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='"
                + fileName.replace("'", "''") + "')?$top=10&$select=id,name";

        Client client = ClientBuilder.newClient();
        try {
            javax.ws.rs.core.Response httpResp = client.target(searchUrl)
                    .request(MediaType.APPLICATION_JSON)
                    .header("Authorization", authHeader)
                    .get();

            if (httpResp.getStatus() != 200) {
                LOG.warn("[MCP-XL] resolveItemId — search returned HTTP {}", httpResp.getStatus());
                return null;
            }

            JsonObject root = GSON.fromJson(httpResp.readEntity(String.class), JsonObject.class);
            JsonArray values = root.has("value") ? root.getAsJsonArray("value") : new JsonArray();

            // Prefer exact name match; fall back to first result
            String fallbackId = null;
            for (JsonElement el : values) {
                if (!el.isJsonObject()) continue;
                JsonObject item = el.getAsJsonObject();
                String itemId   = getStr(item, "id");
                String itemName = getStr(item, "name");
                if (itemId.isEmpty()) continue;
                if (fallbackId == null) fallbackId = itemId;
                if (itemName.equalsIgnoreCase(fileName)) {
                    LOG.info("[MCP-XL] resolveItemId — exact match '{}' → id: {}", fileName, itemId);
                    return path.replace("{itemId}", itemId);
                }
            }

            if (fallbackId != null) {
                LOG.info("[MCP-XL] resolveItemId — first match used for '{}' → id: {}", fileName, fallbackId);
                return path.replace("{itemId}", fallbackId);
            }

            LOG.warn("[MCP-XL] resolveItemId — file '{}' not found in drive", fileName);
            return null;

        } catch (Exception e) {
            LOG.error("[MCP-XL] resolveItemId FAILED — {}", e.getMessage(), e);
            return null;
        } finally {
            client.close();
        }
    }

    /** Extracts a filename (e.g. {@code Book.xlsx}) from the user's question. */
    private static String extractFileName(String question) {
        Matcher m = Pattern.compile("([\\w\\-. ]+\\.(?:xlsx|xls|xlsm|xlsb))",
                Pattern.CASE_INSENSITIVE).matcher(question);
        return m.find() ? m.group(1).trim() : null;
    }

    /**
     * Converts any Graph API JSON response into FreeForm rows.
     *
     * <p>Handles three shapes:
     * <ul>
     *   <li>Search response — {@code value[].hitsContainers[].hits[].resource}</li>
     *   <li>List response  — flat {@code value[]} array (driveItems, etc.)</li>
     *   <li>Single object  — no {@code value} key (e.g. GET /me/drive/items/{id})</li>
     * </ul>
     */
    private static List<FreeForm> parseGenericResponse(String json) {
        List<FreeForm> rows = new ArrayList<>();
        JsonObject root = GSON.fromJson(json, JsonObject.class);

        if (!root.has("value") || !root.get("value").isJsonArray()) {
            rows.add(graphObjectToFreeForm(root));
            return rows;
        }

        JsonArray value = root.getAsJsonArray("value");
        if (value.size() == 0) return rows;

        JsonElement first = value.get(0);
        if (first.isJsonObject() && first.getAsJsonObject().has("hitsContainers")) {
            // Search response
            for (JsonElement valEl : value) {
                JsonArray hitsContainers = valEl.getAsJsonObject().has("hitsContainers")
                        ? valEl.getAsJsonObject().getAsJsonArray("hitsContainers")
                        : new JsonArray();
                for (JsonElement hcEl : hitsContainers) {
                    JsonArray hits = hcEl.getAsJsonObject().has("hits")
                            ? hcEl.getAsJsonObject().getAsJsonArray("hits")
                            : new JsonArray();
                    for (JsonElement hitEl : hits) {
                        JsonObject hit      = hitEl.getAsJsonObject();
                        JsonObject resource = hit.has("resource")
                                ? hit.getAsJsonObject("resource") : new JsonObject();
                        FreeForm ff = graphObjectToFreeForm(resource);
                        if (hit.has("hitId") && !hit.get("hitId").isJsonNull()) {
                            String hitId = hit.get("hitId").getAsString();
                            if (!hitId.isEmpty()) ff.put("FileId", " Text ", hitId);
                        }
                        rows.add(ff);
                    }
                }
            }
            return rows;
        }

        // Flat list response — may be driveItems or sharedWithMe items
        for (JsonElement el : value) {
            if (!el.isJsonObject()) continue;
            JsonObject item = el.getAsJsonObject();
            // sharedWithMe items wrap remote details inside "remoteItem";
            // merge remoteItem fields into the top-level object so name/webUrl/size are found.
            if (item.has("remoteItem") && item.get("remoteItem").isJsonObject()) {
                JsonObject merged = new JsonObject();
                // copy top-level fields first, then overlay with remoteItem fields
                for (Map.Entry<String, JsonElement> e : item.entrySet()) {
                    if (!"remoteItem".equals(e.getKey())) merged.add(e.getKey(), e.getValue());
                }
                for (Map.Entry<String, JsonElement> e : item.getAsJsonObject("remoteItem").entrySet()) {
                    merged.add(e.getKey(), e.getValue());
                }
                rows.add(graphObjectToFreeForm(merged));
            } else {
                rows.add(graphObjectToFreeForm(item));
            }
        }
        return rows;
    }

    /**
     * Converts a single Graph JSON object into a FreeForm row by iterating
     * all its fields and mapping known field names to display-friendly keys.
     */
    // Workbook range metadata fields that are not useful to display
    private static final java.util.Set<String> RANGE_SKIP_FIELDS = java.util.Set.of(
            "columnHidden", "rowHidden", "columnIndex", "rowIndex",
            "cellCount", "columnCount", "rowCount", "hidden",
            "addressLocal", "formulasR1C1", "formulasLocal", "formulasWithA1R1C1",
            "numberFormat", "numberFormatLocal", "valueTypes", "text",
            "formulas", "formulasWithA1R1C1Local");

    private static FreeForm graphObjectToFreeForm(JsonObject obj) {
        FreeForm ff = new FreeForm();
        for (Map.Entry<String, JsonElement> entry : obj.entrySet()) {
            if (entry.getKey().startsWith("@")) continue; // skip @odata.* metadata fields
            if (RANGE_SKIP_FIELDS.contains(entry.getKey())) continue; // skip noisy range metadata
            String strVal = toDisplayString(entry.getKey(), entry.getValue(), obj);
            if (strVal != null && !strVal.isEmpty()) {
                ff.put(toFreeFormKey(entry.getKey()), " Text ", strVal);
            }
        }
        return ff;
    }

    /**
     * Converts a Graph JSON field to a display string.
     * Handles primitives, date truncation, and nested identity/file objects.
     */
    private static String toDisplayString(String key, JsonElement val, JsonObject parent) {
        if (val == null || val.isJsonNull()) return null;
        if (val.isJsonPrimitive()) {
            String s = val.getAsString();
            if ((key.contains("DateTime") || key.contains("date")) && s.contains("T")) {
                return s.substring(0, s.indexOf('T'));
            }
            if ("size".equals(key)) {
                try {
                    long bytes = Long.parseLong(s);
                    if (bytes < 1024) return bytes + " B";
                    if (bytes < 1024 * 1024) return (bytes / 1024) + " KB";
                    return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
                } catch (NumberFormatException e) {
                    return s;
                }
            }
            return s;
        }
        if (val.isJsonObject()) {
            JsonObject obj = val.getAsJsonObject();
            if ("lastModifiedBy".equals(key) || "createdBy".equals(key)) {
                JsonObject user = obj.has("user") && !obj.get("user").isJsonNull()
                        ? obj.getAsJsonObject("user") : null;
                return user != null ? getStr(user, "displayName") : "";
            }
            if ("parentReference".equals(key)) return getStr(obj, "name");
            if ("file".equals(key))            return getStr(obj, "mimeType");
        }
        if (val.isJsonArray()) {
            // Workbook range values/formulas are 2D arrays: [[row1col1, row1col2], [row2col1, ...]]
            if ("values".equals(key)) {
                return flattenValueArray(val.getAsJsonArray());
            }
        }
        return null;
    }

    /**
     * Flattens a 2D cell value array {@code [["r1c1","r1c2"],["r2c1","r2c2"]]}
     * into a display string. Rows are separated by {@code " / "}, columns by {@code " | "}.
     */
    private static String flattenValueArray(JsonArray arr2D) {
        if (arr2D == null || arr2D.size() == 0) return null;
        StringBuilder sb = new StringBuilder();
        for (JsonElement rowEl : arr2D) {
            if (!rowEl.isJsonArray()) continue;
            JsonArray row = rowEl.getAsJsonArray();
            StringBuilder rowSb = new StringBuilder();
            for (JsonElement cell : row) {
                if (rowSb.length() > 0) rowSb.append(" | ");
                rowSb.append(cell.isJsonNull() ? "" : cell.getAsString());
            }
            if (rowSb.length() > 0) {
                if (sb.length() > 0) sb.append(" / ");
                sb.append(rowSb);
            }
        }
        return sb.length() > 0 ? sb.toString() : null;
    }

    /**
     * Maps a Graph driveItem field name to a PascalCase FreeForm key.
     * Unknown fields pass through unchanged.
     */
    private static String toFreeFormKey(String graphKey) {
        switch (graphKey) {
            // driveItem fields
            case "name":                 return "Name";
            case "webUrl":               return "Link";
            case "lastModifiedDateTime": return "LastModified";
            case "createdDateTime":      return "Created";
            case "size":                 return "Size";
            case "id":                   return "FileId";
            case "lastModifiedBy":       return "ModifiedBy";
            case "createdBy":            return "CreatedBy";
            case "parentReference":      return "Folder";
            case "file":                 return "MimeType";
            // workbook range fields
            case "address":              return "Cell";
            case "values":               return "Value";
            default:                     return graphKey;
        }
    }

    // ── Shared helpers ────────────────────────────────────────────────────

    private static String callAi(String apiKey, String provider, String prompt) throws Exception {
        if ("openai".equals(provider)) {
            return OpenAiApiClient.call(apiKey, prompt, McpModelConfig.openAiModel(), 0);
        } else if ("gemini".equals(provider)) {
            return GeminiApiClient.call(apiKey, prompt, McpModelConfig.geminiModel(), 0);
        } else {
            return AnthropicApiClient.call(apiKey, prompt, McpModelConfig.anthropicModel(), 0);
        }
    }

    private static ExtensionResponse success(Map<String, Object> values) {
        return new ExtensionResponse(ExtensionResponse.Result.SUCCESS, values, null, null, null);
    }

    private static ExtensionResponse failure(String message) {
        ExtensionResponse.Error error = new ExtensionResponse.Error(
                message, System.currentTimeMillis(),
                ExtensionResponse.Error.ExceptionType.LOGIC_ERROR, "");
        RemediationActions actions = new RemediationActions(
                List.of(InformAPersonAction.create(message,
                        RemediationAction.RecipientType.ACTIVE_USER, List.of())), null);
        return new ExtensionResponse(ExtensionResponse.Result.FAILURE, null, error, actions, Map.of());
    }

    private static String getJsonString(JsonObject obj, String key) {
        if (obj == null || !obj.has(key) || obj.get(key).isJsonNull()) return "";
        return obj.get(key).getAsString();
    }

    private static String getStr(JsonObject obj, String key) {
        if (obj == null || !obj.has(key) || obj.get(key).isJsonNull()) return "";
        return obj.get(key).getAsString();
    }

    private static boolean blank(String s) {
        return s == null || s.isBlank();
    }

    /**
     * Adapts the Excel {@link RefreshTokenStore} (put/get/remove) to
     * the SDK's {@link ITokenStore} interface (save/get/delete).
     *
     * <p>The Excel RefreshTokenStore methods throw checked {@link IOException},
     * which are wrapped in {@link UncheckedIOException}.</p>
     */
    private static ITokenStore adaptTokenStore(RefreshTokenStore store) {
        return new ITokenStore() {
            @Override
            public void save(String key, String value) {
                try {
                    store.put(key, value);
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }

            @Override
            public String get(String key) {
                try {
                    return store.get(key);
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            }

            @Override
            public void delete(String key) {
                store.remove(key);
            }
        };
    }
}
