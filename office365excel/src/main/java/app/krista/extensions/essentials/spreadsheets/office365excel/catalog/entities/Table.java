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
        ecosystemVersion = "9794e9eb-1731-4077-8b4e-04ef8e1b9641")
@Entity(name = "Table", id = "localDomainEntity_817f8cd8-a960-4992-9893-ea44eb67fa84", primaryKey = "Id", supportStore = false)
public class Table {

    @Field(name = "Id", type = "Text")
    public String id;

    @Field(name = "Name", type = "Text")
    public String name;


    @Field(name = "SheetName", type = "Text")
    public String sheetName;

    public Table(String id, String name, String SheetName) {
        this.id = id;
        this.name = name;
        this.sheetName = SheetName;
    }

    public Table(String id, String name) {
        this(id, name, null);
    }

    public Table() {
    }

}