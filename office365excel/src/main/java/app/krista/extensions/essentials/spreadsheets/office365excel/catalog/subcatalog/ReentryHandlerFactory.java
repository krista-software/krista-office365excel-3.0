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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.handlers.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.impl.*;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service.TelemetryHelper;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.RestClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.*;

import javax.inject.Inject;
import javax.inject.Singleton;

import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import org.jvnet.hk2.annotations.Service;

/**
 * Factory for creating appropriate reentry handlers
 */
@Service
@Singleton
public class ReentryHandlerFactory {

    private final ErrorHandlingStateManager stateManager;
    private final ExtensionResponseGenerator responseGenerator;
    private final WorkbookService workbookService;
    private final WorksheetService worksheetService;
    private final TableService tableService;
    private final RestClientProvider restClientProvider;
    private final DriveService driveService;
    private final WorksheetOperation worksheetOperation;
    private final StringUtil stringUtil;
    private final SheetFormattingService sheetFormattingService;
    private final TelemetryHelper telemetryHelper;

    /**
     * Constructs a new ReentryHandlerFactory with the required dependencies
     *
     * @param stateManager       The error handling state manager for maintaining operation state
     * @param responseGenerator  The response generator for creating extension responses
     * @param workbookService    Service for workbook operations
     * @param worksheetService   Service for worksheet operations
     * @param tableService       Service for table operations
     * @param restClientProvider Provider for REST client instances
     * @param driveService       Service for drive operations
     * @param worksheetOperation Service for worksheet operations
     */
    @Inject
    public ReentryHandlerFactory(ErrorHandlingStateManager stateManager, ExtensionResponseGenerator responseGenerator,
                                 WorkbookService workbookService, WorksheetService worksheetService, TableService tableService,
                                 RestClientProvider restClientProvider, DriveService driveService, WorksheetOperation worksheetOperation,
                                 StringUtil stringUtil, SheetFormattingService sheetFormattingService, TelemetryHelper telemetryHelper) {
        this.stateManager = stateManager;
        this.responseGenerator = responseGenerator;
        this.workbookService = workbookService;
        this.worksheetService = worksheetService;
        this.tableService = tableService;
        this.restClientProvider = restClientProvider;
        this.driveService = driveService;
        this.worksheetOperation = worksheetOperation;
        this.stringUtil = stringUtil;
        this.sheetFormattingService = sheetFormattingService;
        this.telemetryHelper = telemetryHelper;
    }

    /**
     * Get the appropriate handler for table operations
     *
     * @return TableReentryHandler instance
     */
    public TableReentryHandler getTableReentryHandler() {
        return new TableReentryHandlerImpl(stateManager, responseGenerator, workbookService, tableService,
                worksheetService, telemetryHelper);
    }

    /**
     * Get the appropriate handler for sheet operations
     *
     * @return SheetReentryHandler instance
     */
    public SheetReentryHandler getSheetReentryHandler() {
        return new SheetReentryHandlerImpl(stateManager, responseGenerator, worksheetService, restClientProvider,
                stringUtil, worksheetOperation, sheetFormattingService, telemetryHelper);
    }

    /**
     * Get the appropriate handler for range operations
     *
     * @return RangeReentryHandler instance
     */
    public RangeReentryHandler getRangeReentryHandler() {
        return new RangeReentryHandlerImpl(stateManager, responseGenerator, workbookService, worksheetService, telemetryHelper);
    }

    /**
     * Get the appropriate handler for cell operations
     *
     * @return CellReentryHandler instance
     */
    public CellReentryHandler getCellReentryHandler() {
        return new CellReentryHandlerImpl(stateManager, responseGenerator, worksheetService, telemetryHelper);
    }

    /**
     * Get the appropriate handler for shared file operations
     *
     * @return SharedFileReentryHandler instance
     */
    public SharedFileReentryHandler getSharedFileReentryHandler() {
        return new SharedFileReentryHandlerImpl(stateManager, responseGenerator, driveService, telemetryHelper);
    }

    /**
     * Get the appropriate handler for column operations
     *
     * @return ColumnReentryHandler instance
     */
    public ColumnReentryHandler getColumnReentryHandler() {
        return new ColumnReentryHandlerImpl(stateManager, responseGenerator, worksheetOperation, workbookService,
                restClientProvider, telemetryHelper);
    }

    /**
     * Get the appropriate handler for PDF operations
     *
     * @return PdfReentryHandler instance
     */
    public PdfReentryHandler getPdfReentryHandler() {
        return new PdfReentryHandlerImpl(stateManager, responseGenerator, driveService, telemetryHelper);
    }

    /**
     * Get the appropriate handler for worksheet operations
     *
     * @return WorksheetReentryHandler instance
     */
    public WorksheetReentryHandler getWorksheetReentryHandler() {
        return new WorksheetReentryHandlerImpl(stateManager, responseGenerator, worksheetService, restClientProvider,
                stringUtil, sheetFormattingService, telemetryHelper);
    }

    /**
     * Get the appropriate handler for row operations
     *
     * @return RowReentryHandler instance
     */
    public RowReentryHandler getRowReentryHandler() {
        return new RowReentryHandlerImpl(stateManager, responseGenerator, worksheetService,restClientProvider,
                workbookService, telemetryHelper);
    }

    /**
     * Get the find value reentry handler
     *
     * @return FindValueReentryHandler
     */
    public FindValueReentryHandler getFindValueReentryHandler() {
        return new FindValueReentryHandlerImpl(stateManager, responseGenerator, worksheetService, telemetryHelper);
    }

    /**
     * Get the workbook reentry handler
     *
     * @return WorkbookReentryHandler
     */
    public WorkbookReentryHandler getWorkbookReentryHandler() {
        return new WorkbookReentryHandlerImpl(stateManager, responseGenerator, workbookService, telemetryHelper);
    }

    /**
     * Get the formula reentry handler
     *
     * @return FormulaReentryHandler
     */
    public FormulaReentryHandler getFormulaReentryHandler() {
        return new FormulaReentryHandlerImpl(workbookService, stateManager, responseGenerator, telemetryHelper);
    }

    /**
     * Get the summation reentry handler
     *
     * @return SummationReentryHandler
     */
    public SummationReentryHandler getSummationReentryHandler() {
        return new SummationReentryHandlerImpl(stateManager, responseGenerator, workbookService, telemetryHelper);
    }

    /**
     * Get the column values reentry handler
     *
     * @return ColumnValuesReentryHandler
     */
    public ColumnValuesReentryHandler getColumnValuesReentryHandler() {
        return new ColumnValuesReentryHandlerImpl(stateManager, responseGenerator, worksheetOperation, telemetryHelper);
    }

    /**
     * Get the file reentry handler
     *
     * @return FileReentryHandler
     */
    public FileReentryHandler getFileReentryHandler() {
        return new FileReentryHandlerImpl(stateManager, responseGenerator, driveService, telemetryHelper);
    }

    /**
     * Get the min max value reentry handler
     *
     * @return MinMaxValueReentryHandler
     */
    public MinMaxValueReentryHandler getMinMaxValueReentryHandler() {
        return new MinMaxValueReentryHandlerImpl(stateManager, responseGenerator, worksheetService, telemetryHelper);
    }

    /**
     * Get the table VLOOKUP reentry handler
     *
     * @return TableVlookupReentryHandler
     */
    public TableVlookupReentryHandler getTableVlookupReentryHandler() {
        return new TableVlookupReentryHandlerImpl(stateManager, responseGenerator, tableService, telemetryHelper);
    }

    /**
     * Get the table rows reentry handler
     *
     * @return TableRowsReentryHandler
     */
    public TableRowsReentryHandler getTableRowsReentryHandler() {
        return new TableRowsReentryHandlerImpl(tableService, stateManager, responseGenerator, workbookService, telemetryHelper);
    }

    /**
     * Get the delete column reentry handler
     *
     * @return DeleteColumnReentryHandler
     */
    public DeleteColumnReentryHandler getTableColumnReentryHandler() {
        return new DeleteColumnReentryHandlerImpl(stateManager, responseGenerator, worksheetService, telemetryHelper);
    }

    /**
     * Get the table creation reentry handler
     *
     * @return TableCreationReentryHandler
     */
    public TableCreationReentryHandler getTableCreationReentryHandler() {
        return new TableCreationReentryHandlerImpl(stateManager, responseGenerator, tableService, telemetryHelper);
    }
}
