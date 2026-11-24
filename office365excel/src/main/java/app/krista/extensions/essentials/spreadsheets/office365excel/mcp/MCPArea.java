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
import app.krista.extension.impl.anno.Attribute;
import app.krista.extension.impl.anno.CatalogRequest;
import app.krista.extension.impl.anno.Domain;
import app.krista.extension.impl.anno.Field;

import javax.inject.Inject;

/**
 * Catalog entry point for the MCP "Ask Agent" natural language query.
 */
@Domain(id = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24",
        name = "Spreadsheets",
        ecosystemId = "catEntryEcosystem_84b53163-327b-4b1b-8c96-9334d292f9f5",
        ecosystemName = "Essentials",
        ecosystemVersion = "fb35f83b-d270-4925-a896-7b677f04f126")
public class MCPArea {

    private final ExcelMcpQueryService mcpQueryService;

    @Inject
    public MCPArea(ExcelMcpQueryService mcpQueryService) {
        this.mcpQueryService = mcpQueryService;
    }

    @CatalogRequest(
            id = "localDomainRequest_b8f2c4a1-6d93-4e7b-a5c1-3f0e72d18b6a",
            name = "Ask Agent",
            description = "Accepts natural language requests, enables the AI agent to reason and invoke available tools as needed, and returns the generated response.",
            area = "MCP",
            type = CatalogRequest.Type.QUERY_SYSTEM, tool = true)
    @Field.Desc(name = "Response", type = "[ FreeForm ]", required = false, description = "List of Excel workbook files matching the query")
    public ExtensionResponse askAgent(
            @Field(name = "Query", type = "Paragraph", required = true, description = "",
                    attributes = {@Attribute(name = "visualWidth", value = "L")}, options = {}) String query) {
        return mcpQueryService.query(query);
    }
}
