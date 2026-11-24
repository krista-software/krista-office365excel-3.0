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
import app.krista.extension.impl.anno.Searchable;

@Domain(id = "catEntryDomain_91825fb8-54b7-445e-90ea-a0f92b347a24",
        name = "Spreadsheets",
        ecosystemId = "catEntryEcosystem_84b53163-327b-4b1b-8c96-9334d292f9f5",
        ecosystemName = "Essentials",
        ecosystemVersion = "60a6d210-d07b-4b1a-b07f-2abb0c42648c")
@Entity(name = "Drive", id = "localDomainEntity_78674eec-06af-4651-9c80-e67bda24a710", primaryKey = "Drive Id", supportStore = false)
public class Drive {

    @Field(name = "Drive Id", type = "Text")
    public String driveId;

    @Searchable
    @Field(name = "Shared By", type = "Text")
    public String sharedBy;

    @Searchable
    @Field(name = "Created By", type = "Text")
    public String createdBy;

    public Drive(String driveId, String sharedBy, String createdBy) {
        this.driveId = driveId;
        this.sharedBy = sharedBy;
        this.createdBy = createdBy;
    }

}
