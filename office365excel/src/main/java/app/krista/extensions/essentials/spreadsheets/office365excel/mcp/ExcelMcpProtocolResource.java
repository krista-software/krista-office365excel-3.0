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

import app.krista.extension.executor.Invoker;
import app.krista.extensions.essentials.spreadsheets.office365excel.OfficeExcelAttributes;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.OfficeExcelAttributeStore;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.RefreshTokenStore;
import app.krista.sdk.msauth.extension.mcp.DefaultMsAuthBridge;
import app.krista.sdk.msauth.extension.mcp.McpSearchContext;
import app.krista.sdk.msauth.token.ITokenStore;
import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import javax.ws.rs.Consumes;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.client.Client;
import javax.ws.rs.client.ClientBuilder;
import javax.ws.rs.client.Entity;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Standard MCP (Model Context Protocol) JSON-RPC endpoint for Office 365 Excel.
 *
 * <p>Registered at {@code /excel365extn/mcp}. Any MCP-compatible client
 * (Claude Desktop, Cursor, Continue, etc.) can connect to this endpoint
 * and invoke Excel tools directly — no Krista UI required.</p>
 *
 * <p>Implements the three core MCP methods:</p>
 * <ul>
 *   <li>{@code initialize}  — handshake / capability negotiation</li>
 *   <li>{@code tools/list} — returns the list of available Excel tools</li>
 *   <li>{@code tools/call} — executes a tool against Microsoft Graph API</li>
 * </ul>
 *
 * <p>Available tools:</p>
 * <ul>
 *   <li>{@code list_workbooks}  — list all Excel files in OneDrive/SharePoint</li>
 *   <li>{@code search_workbooks} — search Excel files by keyword</li>
 *   <li>{@code get_workbook}    — get details of a specific workbook by file ID</li>
 *   <li>{@code list_worksheets} — list sheets inside a workbook</li>
 *   <li>{@code read_worksheet}  — read cell data from a worksheet</li>
 * </ul>
 */
@Path("/mcp")
public class ExcelMcpProtocolResource {

    private static final Logger LOG = LoggerFactory.getLogger(ExcelMcpProtocolResource.class);
    private static final Gson GSON = new Gson();
    private static final String GRAPH_BASE = "https://graph.microsoft.com/v1.0";
    private static final String JSONRPC = "2.0";

    private final ITokenStore tokenStore;
    private final Invoker invoker;
    private final OfficeExcelAttributeStore attributeStore;

    @Inject
    public ExcelMcpProtocolResource(RefreshTokenStore refreshTokenStore,
                                    Invoker invoker,
                                    OfficeExcelAttributeStore attributeStore) {
        this.tokenStore   = adaptTokenStore(refreshTokenStore);
        this.invoker      = invoker;
        this.attributeStore = attributeStore;
    }

    // ── Main dispatcher ───────────────────────────────────────────────────

    @POST
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    public Response handle(String requestBody) {
        LOG.debug("[MCP-XL-PROTO] incoming request: {}", requestBody);

        JsonObject req;
        try {
            req = GSON.fromJson(requestBody, JsonObject.class);
        } catch (Exception e) {
            return mcpError(null, -32700, "Parse error: " + e.getMessage());
        }

        String method  = req.has("method") ? req.get("method").getAsString() : "";
        JsonElement id = req.has("id") ? req.get("id") : null; // null = notification

        try {
            switch (method) {

                case "initialize":
                    return ok(id, buildInitializeResult());

                case "notifications/initialized":
                    // Notification — client acknowledges ready; no response body needed
                    return Response.ok("{}").type(MediaType.APPLICATION_JSON).build();

                case "ping":
                    return ok(id, new JsonObject());

                case "tools/list":
                    return ok(id, buildToolsList());

                case "tools/call":
                    return handleToolsCall(id, req);

                default:
                    LOG.warn("[MCP-XL-PROTO] unknown method: {}", method);
                    return mcpError(id, -32601, "Method not found: " + method);
            }
        } catch (Exception e) {
            LOG.error("[MCP-XL-PROTO] internal error handling '{}': {}", method, e.getMessage(), e);
            return mcpError(id, -32603, "Internal error: " + e.getMessage());
        }
    }

    // ── initialize ────────────────────────────────────────────────────────

    private static JsonObject buildInitializeResult() {
        JsonObject serverInfo = new JsonObject();
        serverInfo.addProperty("name", "krista-excel-mcp");
        serverInfo.addProperty("version", "1.0.0");

        JsonObject capabilities = new JsonObject();
        capabilities.add("tools", new JsonObject());

        JsonObject result = new JsonObject();
        result.addProperty("protocolVersion", "2024-11-05");
        result.add("serverInfo", serverInfo);
        result.add("capabilities", capabilities);
        return result;
    }

    // ── tools/list ────────────────────────────────────────────────────────

    private static JsonObject buildToolsList() {
        JsonArray tools = new JsonArray();
        tools.add(buildTool(
                "list_workbooks",
                "List all Excel workbooks (xlsx, xls, xlsm) stored in OneDrive and SharePoint.",
                buildSchema(List.of(), Map.of())
        ));
        tools.add(buildTool(
                "search_workbooks",
                "Search for Excel workbooks by keyword (title, author, content).",
                buildSchema(
                        List.of("query"),
                        Map.of("query", strProp("Search keyword, e.g. 'budget' or 'sales report'"))
                )
        ));
        tools.add(buildTool(
                "get_workbook",
                "Get full details of a specific Excel workbook using its file ID.",
                buildSchema(
                        List.of("fileId"),
                        Map.of("fileId", strProp("The OneDrive file ID of the workbook"))
                )
        ));
        tools.add(buildTool(
                "list_worksheets",
                "List all worksheets (tabs) inside an Excel workbook.",
                buildSchema(
                        List.of("fileId"),
                        Map.of("fileId", strProp("The OneDrive file ID of the workbook"))
                )
        ));
        tools.add(buildTool(
                "read_worksheet",
                "Read the cell data from a worksheet in an Excel workbook.",
                buildSchema(
                        List.of("fileId"),
                        Map.of(
                                "fileId", strProp("The OneDrive file ID of the workbook"),
                                "sheet",  strProp("Worksheet name, e.g. 'Sheet1'. Defaults to 'Sheet1' if omitted.")
                        )
                )
        ));

        JsonObject result = new JsonObject();
        result.add("tools", tools);
        return result;
    }

    // ── tools/call ────────────────────────────────────────────────────────

    private Response handleToolsCall(JsonElement id, JsonObject req) {
        JsonObject params    = getObj(req, "params");
        String toolName      = getStr(params, "name");
        JsonObject arguments = params.has("arguments") && params.get("arguments").isJsonObject()
                ? params.getAsJsonObject("arguments") : new JsonObject();

        LOG.info("[MCP-XL-PROTO] tools/call — tool: '{}', args: {}", toolName, arguments);

        // Authenticate
        DefaultMsAuthBridge authBridge = currentAuthBridge();
        if (!authBridge.isAuthenticated()) {
            return toolError(id, "Not connected. Please configure credentials in the "
                    + "Krista Authentication tab for Office 365 Excel.");
        }

        String authHeader;
        try {
            McpSearchContext ctx = authBridge.buildSearchContext(new ExcelMcpProductConfig());
            authHeader = ctx.authHeader();
        } catch (Exception e) {
            LOG.error("[MCP-XL-PROTO] auth failed: {}", e.getMessage(), e);
            return toolError(id, "Authentication failed: " + e.getMessage());
        }

        // Route to tool handler
        switch (toolName) {
            case "list_workbooks":
                return toolResult(id, execListWorkbooks(authHeader));

            case "search_workbooks": {
                String query = getStr(arguments, "query");
                if (query.isEmpty()) return toolError(id, "'query' argument is required.");
                return toolResult(id, execSearchWorkbooks(authHeader, query));
            }

            case "get_workbook": {
                String fileId = getStr(arguments, "fileId");
                if (fileId.isEmpty()) return toolError(id, "'fileId' argument is required.");
                return toolResult(id, execGetWorkbook(authHeader, fileId));
            }

            case "list_worksheets": {
                String fileId = getStr(arguments, "fileId");
                if (fileId.isEmpty()) return toolError(id, "'fileId' argument is required.");
                return toolResult(id, execListWorksheets(authHeader, fileId));
            }

            case "read_worksheet": {
                String fileId = getStr(arguments, "fileId");
                if (fileId.isEmpty()) return toolError(id, "'fileId' argument is required.");
                String sheet = getStr(arguments, "sheet");
                if (sheet.isEmpty()) sheet = "Sheet1";
                return toolResult(id, execReadWorksheet(authHeader, fileId, sheet));
            }

            default:
                return mcpError(id, -32602, "Unknown tool: " + toolName);
        }
    }

    // ── Tool implementations ──────────────────────────────────────────────

    /**
     * list_workbooks — searches OneDrive/SharePoint for all Excel file types.
     * Uses the Graph Search API so it covers both own and shared workbooks.
     */
    private String execListWorkbooks(String authHeader) {
        String body = buildSearchBody("FileType:xlsx OR FileType:xls OR FileType:xlsm OR FileType:xlsb", 50);
        String json = graphPost(GRAPH_BASE + "/search/query", body, authHeader);
        if (json == null) return "Error: failed to reach Microsoft Graph API.";
        return formatSearchResponse(json, "workbooks");
    }

    /**
     * search_workbooks — searches by keyword, scoped to Excel file types.
     */
    private String execSearchWorkbooks(String authHeader, String keyword) {
        String scopedQuery = keyword + " FileType:xlsx OR FileType:xls OR FileType:xlsm";
        String body = buildSearchBody(scopedQuery, 50);
        String json = graphPost(GRAPH_BASE + "/search/query", body, authHeader);
        if (json == null) return "Error: failed to reach Microsoft Graph API.";
        return formatSearchResponse(json, "workbooks matching \"" + keyword + "\"");
    }

    /**
     * get_workbook — fetches a single driveItem by file ID.
     */
    private String execGetWorkbook(String authHeader, String fileId) {
        String url = GRAPH_BASE + resolvedPath("/me/drive/items/" + fileId);
        String json = graphGet(url, authHeader);
        if (json == null) return "Error: failed to reach Microsoft Graph API.";
        return formatSingleItem(json);
    }

    /**
     * list_worksheets — lists all sheet tabs in a workbook.
     */
    private String execListWorksheets(String authHeader, String fileId) {
        String url = GRAPH_BASE + resolvedPath("/me/drive/items/" + fileId + "/workbook/worksheets");
        String json = graphGet(url, authHeader);
        if (json == null) return "Error: failed to reach Microsoft Graph API.";
        return formatWorksheetsResponse(json);
    }

    /**
     * read_worksheet — reads the used range of a named worksheet.
     */
    private String execReadWorksheet(String authHeader, String fileId, String sheet) {
        String encoded = sheet.replace(" ", "%20");
        String url = GRAPH_BASE + resolvedPath(
                "/me/drive/items/" + fileId + "/workbook/worksheets/" + encoded + "/usedRange");
        String json = graphGet(url, authHeader);
        if (json == null) return "Error: failed to reach Microsoft Graph API.";
        return formatWorksheetData(json, sheet);
    }

    // ── Graph API response formatters ─────────────────────────────────────

    /**
     * Formats a Graph Search API response (hitsContainers) into readable text.
     */
    private static String formatSearchResponse(String json, String label) {
        try {
            JsonObject root = GSON.fromJson(json, JsonObject.class);

            // Check for Graph error
            if (root.has("error")) {
                return "Graph API error: " + getStr(root.getAsJsonObject("error"), "message");
            }

            List<Map<String, String>> items = new ArrayList<>();
            JsonArray values = root.has("value") ? root.getAsJsonArray("value") : new JsonArray();

            for (JsonElement valEl : values) {
                JsonArray hitsContainers = valEl.getAsJsonObject().has("hitsContainers")
                        ? valEl.getAsJsonObject().getAsJsonArray("hitsContainers") : new JsonArray();
                for (JsonElement hcEl : hitsContainers) {
                    JsonArray hits = hcEl.getAsJsonObject().has("hits")
                            ? hcEl.getAsJsonObject().getAsJsonArray("hits") : new JsonArray();
                    for (JsonElement hitEl : hits) {
                        JsonObject hit      = hitEl.getAsJsonObject();
                        JsonObject resource = hit.has("resource")
                                ? hit.getAsJsonObject("resource") : new JsonObject();
                        Map<String, String> item = new LinkedHashMap<>();
                        item.put("name",       getStr(resource, "name"));
                        item.put("modified",   truncateDate(getStr(resource, "lastModifiedDateTime")));
                        item.put("modifiedBy", extractDisplayName(resource, "lastModifiedBy"));
                        item.put("size",       formatSize(getStr(resource, "size")));
                        item.put("link",       getStr(resource, "webUrl"));
                        item.put("fileId",     hit.has("hitId") ? getStr(hit, "hitId") : getStr(resource, "id"));
                        items.add(item);
                    }
                }
            }

            if (items.isEmpty()) return "No " + label + " found.";

            StringBuilder sb = new StringBuilder();
            sb.append("Found ").append(items.size()).append(" ").append(label).append(":\n");
            for (int i = 0; i < items.size(); i++) {
                Map<String, String> it = items.get(i);
                sb.append("\n").append(i + 1).append(". ").append(it.get("name")).append("\n");
                if (!it.get("modified").isEmpty())
                    sb.append("   Modified:    ").append(it.get("modified")).append("\n");
                if (!it.get("modifiedBy").isEmpty())
                    sb.append("   Modified by: ").append(it.get("modifiedBy")).append("\n");
                if (!it.get("size").isEmpty())
                    sb.append("   Size:        ").append(it.get("size")).append("\n");
                if (!it.get("link").isEmpty())
                    sb.append("   Link:        ").append(it.get("link")).append("\n");
                if (!it.get("fileId").isEmpty())
                    sb.append("   File ID:     ").append(it.get("fileId")).append("\n");
            }
            return sb.toString().trim();

        } catch (Exception e) {
            LOG.error("[MCP-XL-PROTO] formatSearchResponse failed: {}", e.getMessage(), e);
            return "Error parsing response: " + e.getMessage();
        }
    }

    /**
     * Formats a single driveItem JSON object into readable text.
     */
    private static String formatSingleItem(String json) {
        try {
            JsonObject obj = GSON.fromJson(json, JsonObject.class);
            if (obj.has("error")) {
                return "Graph API error: " + getStr(obj.getAsJsonObject("error"), "message");
            }
            StringBuilder sb = new StringBuilder();
            sb.append("Workbook: ").append(getStr(obj, "name")).append("\n");
            sb.append("File ID:  ").append(getStr(obj, "id")).append("\n");
            String size = formatSize(getStr(obj, "size"));
            if (!size.isEmpty()) sb.append("Size:     ").append(size).append("\n");
            String modified = truncateDate(getStr(obj, "lastModifiedDateTime"));
            if (!modified.isEmpty()) sb.append("Modified: ").append(modified).append("\n");
            String modBy = extractDisplayName(obj, "lastModifiedBy");
            if (!modBy.isEmpty()) sb.append("Modified by: ").append(modBy).append("\n");
            String link = getStr(obj, "webUrl");
            if (!link.isEmpty()) sb.append("Link:     ").append(link).append("\n");
            return sb.toString().trim();
        } catch (Exception e) {
            return "Error parsing response: " + e.getMessage();
        }
    }

    /**
     * Formats a worksheets list response into readable text.
     */
    private static String formatWorksheetsResponse(String json) {
        try {
            JsonObject root = GSON.fromJson(json, JsonObject.class);
            if (root.has("error")) {
                return "Graph API error: " + getStr(root.getAsJsonObject("error"), "message");
            }
            JsonArray value = root.has("value") ? root.getAsJsonArray("value") : new JsonArray();
            if (value.size() == 0) return "No worksheets found in this workbook.";

            StringBuilder sb = new StringBuilder();
            sb.append(value.size()).append(" worksheet(s):\n");
            for (int i = 0; i < value.size(); i++) {
                JsonObject ws = value.get(i).getAsJsonObject();
                String name       = getStr(ws, "name");
                String visibility = getStr(ws, "visibility");
                sb.append("\n").append(i + 1).append(". ").append(name);
                if (!visibility.isEmpty() && !"Visible".equals(visibility)) {
                    sb.append(" (").append(visibility.toLowerCase()).append(")");
                }
                sb.append("\n");
            }
            return sb.toString().trim();
        } catch (Exception e) {
            return "Error parsing response: " + e.getMessage();
        }
    }

    /**
     * Formats a usedRange response (2-D values array) as a plain-text table.
     */
    private static String formatWorksheetData(String json, String sheetName) {
        try {
            JsonObject root = GSON.fromJson(json, JsonObject.class);
            if (root.has("error")) {
                return "Graph API error: " + getStr(root.getAsJsonObject("error"), "message");
            }

            JsonArray rows = root.has("values") ? root.getAsJsonArray("values") : new JsonArray();
            if (rows.size() == 0) return "Sheet \"" + sheetName + "\" is empty.";

            int rowCount = rows.size();
            int colCount = rows.get(0).getAsJsonArray().size();

            // Calculate column widths for alignment
            int[] widths = new int[colCount];
            List<List<String>> table = new ArrayList<>();
            for (JsonElement rowEl : rows) {
                JsonArray rowArr = rowEl.getAsJsonArray();
                List<String> row = new ArrayList<>();
                for (int c = 0; c < colCount; c++) {
                    String cell = c < rowArr.size() ? rowArr.get(c).getAsString() : "";
                    row.add(cell);
                    widths[c] = Math.max(widths[c], Math.min(cell.length(), 30));
                }
                table.add(row);
            }

            StringBuilder sb = new StringBuilder();
            sb.append("Sheet: \"").append(sheetName).append("\" — ")
              .append(rowCount).append(" rows × ").append(colCount).append(" columns\n\n");

            // Header row
            List<String> header = table.get(0);
            appendRow(sb, header, widths);
            appendSeparator(sb, widths);

            // Data rows (limit to 100 to avoid enormous output)
            int limit = Math.min(table.size(), 101);
            for (int r = 1; r < limit; r++) {
                appendRow(sb, table.get(r), widths);
            }
            if (table.size() > 101) {
                sb.append("\n... and ").append(table.size() - 101).append(" more rows.");
            }
            return sb.toString().trim();

        } catch (Exception e) {
            return "Error parsing worksheet data: " + e.getMessage();
        }
    }

    // ── Table formatting helpers ──────────────────────────────────────────

    private static void appendRow(StringBuilder sb, List<String> row, int[] widths) {
        for (int c = 0; c < widths.length; c++) {
            String cell = c < row.size() ? row.get(c) : "";
            if (cell.length() > 30) cell = cell.substring(0, 27) + "...";
            sb.append(pad(cell, widths[c]));
            if (c < widths.length - 1) sb.append(" | ");
        }
        sb.append("\n");
    }

    private static void appendSeparator(StringBuilder sb, int[] widths) {
        for (int c = 0; c < widths.length; c++) {
            sb.append("-".repeat(widths[c]));
            if (c < widths.length - 1) sb.append("-+-");
        }
        sb.append("\n");
    }

    private static String pad(String s, int width) {
        if (s.length() >= width) return s;
        return s + " ".repeat(width - s.length());
    }

    // ── HTTP helpers ──────────────────────────────────────────────────────

    private static String graphPost(String url, String body, String authHeader) {
        Client client = ClientBuilder.newClient();
        try {
            javax.ws.rs.core.Response resp = client.target(url)
                    .request(MediaType.APPLICATION_JSON)
                    .header("Authorization", authHeader)
                    .post(Entity.entity(body, MediaType.APPLICATION_JSON));
            if (resp.getStatus() == 200 || resp.getStatus() == 201) {
                return resp.readEntity(String.class);
            }
            String err = resp.readEntity(String.class);
            LOG.error("[MCP-XL-PROTO] POST {} → HTTP {}: {}", url, resp.getStatus(), err);
            return wrapGraphError(resp.getStatus(), err);
        } catch (Exception e) {
            LOG.error("[MCP-XL-PROTO] POST {} failed: {}", url, e.getMessage(), e);
            return null;
        } finally {
            client.close();
        }
    }

    private static String graphGet(String url, String authHeader) {
        Client client = ClientBuilder.newClient();
        try {
            javax.ws.rs.core.Response resp = client.target(url)
                    .request(MediaType.APPLICATION_JSON)
                    .header("Authorization", authHeader)
                    .get();
            if (resp.getStatus() == 200) {
                return resp.readEntity(String.class);
            }
            String err = resp.readEntity(String.class);
            LOG.error("[MCP-XL-PROTO] GET {} → HTTP {}: {}", url, resp.getStatus(), err);
            return wrapGraphError(resp.getStatus(), err);
        } catch (Exception e) {
            LOG.error("[MCP-XL-PROTO] GET {} failed: {}", url, e.getMessage(), e);
            return null;
        } finally {
            client.close();
        }
    }

    /**
     * Wraps a non-200 Graph response body into a JSON object with an {@code error}
     * field so that the formatters can detect and surface it cleanly.
     */
    private static String wrapGraphError(int status, String body) {
        try {
            // If Graph already returned an error object, pass it through
            JsonObject parsed = GSON.fromJson(body, JsonObject.class);
            if (parsed.has("error")) return body;
        } catch (Exception ignored) { }
        JsonObject wrapper = new JsonObject();
        JsonObject error   = new JsonObject();
        error.addProperty("code", String.valueOf(status));
        error.addProperty("message", "HTTP " + status + ": " + body);
        wrapper.add("error", error);
        return wrapper.toString();
    }

    // ── Auth helpers ──────────────────────────────────────────────────────

    private DefaultMsAuthBridge currentAuthBridge() {
        OfficeExcelAttributes attrs = attributeStore.load(invoker.getInvokerId());
        return new DefaultMsAuthBridge(attrs, tokenStore);
    }

    /**
     * Replaces {@code /me/} with {@code /users/{email}/} when Private
     * (client-credentials) auth is configured.
     */
    private String resolvedPath(String path) {
        OfficeExcelAttributes attrs = attributeStore.load(invoker.getInvokerId());
        if (attrs != null && "PRIVATE".equals(attrs.getAuthType())) {
            String email = attrs.getEmail();
            if (email != null && !email.isBlank()) {
                return path.replaceFirst("^/me(/|$)", "/users/" + email + "$1");
            }
        }
        return path;
    }

    // ── JSON-RPC response builders ────────────────────────────────────────

    private static Response ok(JsonElement id, JsonObject result) {
        JsonObject resp = new JsonObject();
        resp.addProperty("jsonrpc", JSONRPC);
        if (id != null) resp.add("id", id);
        resp.add("result", result);
        return Response.ok(GSON.toJson(resp)).type(MediaType.APPLICATION_JSON).build();
    }

    /** Returns a successful tool result with text content. */
    private static Response toolResult(JsonElement id, String text) {
        JsonObject content = new JsonObject();
        content.addProperty("type", "text");
        content.addProperty("text", text);

        JsonArray contentArr = new JsonArray();
        contentArr.add(content);

        JsonObject result = new JsonObject();
        result.add("content", contentArr);

        return ok(id, result);
    }

    /**
     * Returns a tool error — note this is a successful JSON-RPC response
     * with {@code isError: true} inside the result, per the MCP spec.
     * The LLM sees the error message and can tell the user what went wrong.
     */
    private static Response toolError(JsonElement id, String message) {
        JsonObject content = new JsonObject();
        content.addProperty("type", "text");
        content.addProperty("text", "Error: " + message);

        JsonArray contentArr = new JsonArray();
        contentArr.add(content);

        JsonObject result = new JsonObject();
        result.add("content", contentArr);
        result.addProperty("isError", true);

        return ok(id, result);
    }

    /** Returns a JSON-RPC protocol-level error (method not found, parse error, etc.). */
    private static Response mcpError(JsonElement id, int code, String message) {
        JsonObject error = new JsonObject();
        error.addProperty("code", code);
        error.addProperty("message", message);

        JsonObject resp = new JsonObject();
        resp.addProperty("jsonrpc", JSONRPC);
        if (id != null) resp.add("id", id);
        resp.add("error", error);

        return Response.status(200).entity(GSON.toJson(resp))
                .type(MediaType.APPLICATION_JSON).build();
    }

    // ── Tool / schema builders ────────────────────────────────────────────

    private static JsonObject buildTool(String name, String description, JsonObject inputSchema) {
        JsonObject tool = new JsonObject();
        tool.addProperty("name", name);
        tool.addProperty("description", description);
        tool.add("inputSchema", inputSchema);
        return tool;
    }

    private static JsonObject buildSchema(List<String> required, Map<String, JsonObject> props) {
        JsonObject schema = new JsonObject();
        schema.addProperty("type", "object");

        JsonObject properties = new JsonObject();
        props.forEach(properties::add);
        schema.add("properties", properties);

        if (!required.isEmpty()) {
            JsonArray req = new JsonArray();
            required.forEach(req::add);
            schema.add("required", req);
        }
        return schema;
    }

    private static JsonObject strProp(String description) {
        JsonObject p = new JsonObject();
        p.addProperty("type", "string");
        p.addProperty("description", description);
        return p;
    }

    // ── Graph request body builders ───────────────────────────────────────

    private static String buildSearchBody(String queryString, int size) {
        JsonObject query = new JsonObject();
        query.addProperty("queryString", queryString);

        JsonArray entityTypes = new JsonArray();
        entityTypes.add("driveItem");

        JsonObject request = new JsonObject();
        request.add("entityTypes", entityTypes);
        request.add("query", query);
        request.addProperty("from", 0);
        request.addProperty("size", size);

        JsonArray requests = new JsonArray();
        requests.add(request);

        JsonObject body = new JsonObject();
        body.add("requests", requests);
        return body.toString();
    }

    // ── Field helpers ─────────────────────────────────────────────────────

    private static String extractDisplayName(JsonObject parent, String fieldName) {
        if (!parent.has(fieldName) || parent.get(fieldName).isJsonNull()) return "";
        JsonObject identitySet = parent.getAsJsonObject(fieldName);
        if (!identitySet.has("user") || identitySet.get("user").isJsonNull()) return "";
        return getStr(identitySet.getAsJsonObject("user"), "displayName");
    }

    private static String truncateDate(String dateTime) {
        if (dateTime == null || dateTime.isEmpty()) return "";
        int tIdx = dateTime.indexOf('T');
        return tIdx > 0 ? dateTime.substring(0, tIdx) : dateTime;
    }

    private static String formatSize(String sizeStr) {
        if (sizeStr == null || sizeStr.isEmpty()) return "";
        try {
            long bytes = Long.parseLong(sizeStr);
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024) + " KB";
            return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
        } catch (NumberFormatException e) {
            return sizeStr;
        }
    }

    private static String getStr(JsonObject obj, String key) {
        if (obj == null || !obj.has(key) || obj.get(key).isJsonNull()) return "";
        return obj.get(key).getAsString();
    }

    private static JsonObject getObj(JsonObject obj, String key) {
        if (obj == null || !obj.has(key) || !obj.get(key).isJsonObject()) return new JsonObject();
        return obj.getAsJsonObject(key);
    }

    // ── Token store adapter ───────────────────────────────────────────────

    private static ITokenStore adaptTokenStore(RefreshTokenStore store) {
        return new ITokenStore() {
            @Override
            public void save(String key, String value) {
                try { store.put(key, value); }
                catch (IOException e) { throw new UncheckedIOException(e); }
            }
            @Override
            public String get(String key) {
                try { return store.get(key); }
                catch (IOException e) { throw new UncheckedIOException(e); }
            }
            @Override
            public void delete(String key) { store.remove(key); }
        };
    }
}
