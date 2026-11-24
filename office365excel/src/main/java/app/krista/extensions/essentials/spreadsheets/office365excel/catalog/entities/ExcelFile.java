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
        ecosystemVersion = "4a25c7a2-ce59-426b-bda4-e396aca1d8ee")
@Entity(name = "Excel File", id = "localDomainEntity_c9b211bf-2ca9-4306-89be-9f69ef112509", primaryKey = "File Id", supportStore = false)
public final class ExcelFile {

    @Field(name = "File Id", type = "Text")
    public String fileId;

    @Field(name = "Drive Id", type = "Text")
    public String driveId;

    @Searchable
    @Field(name = "File Name", type = "Text")
    public String fileName;

    @Field(name = "Web URL", type = "Text")
    public String webURL;

    public ExcelFile(String fileId, String driveId, String fileName, String webURL) {
        this.fileId = fileId;
        this.driveId = driveId;
        this.fileName = fileName;
        this.webURL = webURL;
    }

}
