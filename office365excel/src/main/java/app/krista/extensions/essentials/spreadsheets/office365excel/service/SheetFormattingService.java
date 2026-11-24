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

package app.krista.extensions.essentials.spreadsheets.office365excel.service;

import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@Service
public class SheetFormattingService {

    private static final Logger LOGGER = LoggerFactory.getLogger(SheetFormattingService.class);

    private final RestClientProvider restClientProvider;

    @Inject
    public SheetFormattingService(RestClientProvider restClientProvider) {
        this.restClientProvider = restClientProvider;
    }

    public void clearRowFormatting(String driveId, String fileId,
                                          String sheetName, Double rowIndex) throws IOException {
        String address = sheetName + "!A" + rowIndex.intValue() + ":Z100000";
        String url = fileId + Constants.WORKBOOK_WORKSHEETS + sheetName +
                "/range(address='" + address + "')/clear";

        LOGGER.debug("[clearRowFormatting] Single-row version ---");
        LOGGER.debug("Sheet: {}", sheetName);
        LOGGER.debug("Row index: {}", rowIndex.intValue());
        LOGGER.debug("Range address to clear: {}", address);
        LOGGER.debug("POST URL: {}", url);

        Map<String, Object> clearBody = Map.of("applyTo", "formats");
        LOGGER.debug("Request body: {}", clearBody);

        restClientProvider.getRestClient(driveId).post(url, clearBody);

        LOGGER.debug("[clearRowFormatting] --- DONE");
    }

    public void clearRowFormatting(String fileId, String sheetName,
                                          double startRowIndex, int numberOfRows) throws IOException {
        int startRow = (int) startRowIndex;
        int endRow = startRow + numberOfRows - 1;

        LOGGER.debug("[clearRowFormatting] Multiple-rows version ---");
        LOGGER.debug("Sheet: {}", sheetName);
        LOGGER.debug("Start row: {}", startRow);
        LOGGER.debug("Number of rows: {}", numberOfRows);
        LOGGER.debug("Calculated end row: {}", endRow);

        String address = sheetName + "!A" + startRow + ":Z" + endRow;
        LOGGER.debug("Range address to clear: {}", address);

        String url = fileId + Constants.WORKBOOK_WORKSHEETS + sheetName + "/range(address='" + address + "')/clear";
        LOGGER.debug("POST URL: {}", url);

        Map<String, Object> clearBody = new HashMap<>();
        clearBody.put("applyTo", "formats");
        LOGGER.debug("Request body: {}", clearBody);

        restClientProvider.getRestClient(null).post(url, clearBody);
        LOGGER.debug("[clearRowFormatting] --- DONE");
    }

}
