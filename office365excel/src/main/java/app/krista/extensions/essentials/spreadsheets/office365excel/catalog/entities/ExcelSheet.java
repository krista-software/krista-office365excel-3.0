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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities;

import app.krista.extension.impl.anno.Domain;
import app.krista.extension.impl.anno.Entity;
import app.krista.extension.impl.anno.Field;

@Domain(id = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24",
        name = "Spreadsheets",
        ecosystemId = "catEntryEcosystem_84b53163-327b-4b1b-8c96-9334d292f9f5",
        ecosystemName = "Essentials",
        ecosystemVersion = "b75c57d8-2fa0-4df2-a491-2bbe57132504")
@Entity(name = "Excel Sheet", id = "localDomainEntity_5ba7f92d-cc40-4d60-b891-dc1ba1d44d90", primaryKey = "Sheet Id", supportStore = false)
public class ExcelSheet {

    @Field(name = "Sheet Id", type = "Text")
    String sheetId;

    @Field(name = "Sheet name", type = "Text")
    String sheetName;

    public ExcelSheet(String sheetId, String sheetName) {
        this.sheetId = sheetId;
        this.sheetName = sheetName;
    }

    public String getSheetName() {
        return sheetName;
    }
}
