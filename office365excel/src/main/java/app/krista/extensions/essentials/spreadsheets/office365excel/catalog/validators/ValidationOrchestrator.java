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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators;

import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorkbookService;
import app.krista.extensions.essentials.spreadsheets.office365excel.service.WorksheetService;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Orchestrates validation across multiple validators.
 * This class is responsible for managing all validators and coordinating validation operations.
 */
@Service
public class ValidationOrchestrator {

    private static final Logger LOGGER = LoggerFactory.getLogger(ValidationOrchestrator.class);
    private final Map<Validator.ValidationResource, Validator> validators = new HashMap<>();

    @Inject
    public ValidationOrchestrator(WorkbookService workbookService, WorksheetService worksheetService) {
        // Initialize all validators
        validators.put(Validator.ValidationResource.DRIVE_ID, new DriveIdValidator(workbookService));
        validators.put(Validator.ValidationResource.FILE_ID, new FileIdValidator(workbookService));
        validators.put(Validator.ValidationResource.TABLE_NAME, new TableNameValidator(workbookService));
        validators.put(Validator.ValidationResource.SHEET_NAME, new SheetNameValidator(worksheetService));
        validators.put(Validator.ValidationResource.ROW_DATA, new RowDataValidator());
        validators.put(Validator.ValidationResource.CSV_FILE, new CSVFileValidator());
        validators.put(Validator.ValidationResource.RANGE, new RangeValidator());
        validators.put(Validator.ValidationResource.COLUMN_INDEX, new ColumnIndexValidator());
        validators.put(Validator.ValidationResource.CELL_ID, new CellIdValidator());
        validators.put(Validator.ValidationResource.ROW_INDEX, new RowIndexValidator());
        validators.put(Validator.ValidationResource.COLUMN_NAME, new ColumnNameValidator(workbookService));
        validators.put(Validator.ValidationResource.KEY_COLUMN_NAME, new ColumnNameValidator(workbookService));
        validators.put(Validator.ValidationResource.FORMULA, new FormulaValidator());
        validators.put(Validator.ValidationResource.FORMULA_NAME, new FormulaNameValidator());
        validators.put(Validator.ValidationResource.COMMENT, new CommentValidator());
        validators.put(Validator.ValidationResource.SHARED_FILE_NAME, new FileNameValidator());
        validators.put(Validator.ValidationResource.KEY_ROWS, new KeyRowsValidator());
        validators.put(Validator.ValidationResource.LOOKUP_COLUMNS, new LookupColumnsValidator());
        validators.put(Validator.ValidationResource.COLUMNS, new ColumnsValidator(workbookService));
        validators.put(Validator.ValidationResource.LOOKUP_ROWS, new LookupRowsValidator());
        validators.put(Validator.ValidationResource.LOOKUP_VALUE, new LookupValueValidator());
        validators.put(Validator.ValidationResource.PATH, new PathValidator());
        validators.put(Validator.ValidationResource.NAME, new FileNameValidator());
        validators.put(Validator.ValidationResource.EXCEL_FILE, new ExcelFileValidator());
        validators.put(Validator.ValidationResource.COLUMN_RANGE, new ColumnRangeValidator());
        validators.put(Validator.ValidationResource.WORKBOOK_ID, new FileIdValidator(workbookService));

    }

    /**
     * Represents the result of a validation operation.
     */
    public static class ValidationResult {
        private final String confirmStepMessage;
        private final String fetchFieldName;
        private final String fetchStepMessage;
        private final String errMessage;
        private final String fieldType;

        public ValidationResult(String confirmStepMessage, String fetchFieldName,
                                String fetchStepMessage, String errMessage, String fieldType) {
            this.confirmStepMessage = confirmStepMessage;
            this.fetchFieldName = fetchFieldName;
            this.fetchStepMessage = fetchStepMessage;
            this.errMessage = errMessage;
            this.fieldType = fieldType;
        }

        // Getters
        public String getConfirmStepMessage() {
            return confirmStepMessage;
        }

        public String getErrMessage() {
            return errMessage;
        }

        public String getFetchFieldName() {
            return fetchFieldName;
        }

        public String getFieldType() {
            return fieldType;
        }

        public String getFetchStepMessage() {
            return fetchStepMessage;
        }
    }

    /**
     * Validates a map of resources against their respective validators
     *
     * @param resources Map of resource types to resource values
     * @return List of validation results for invalid resources
     */
    public List<ValidationResult> validate(Map<Validator.ValidationResource, String> resources) {
        List<ValidationResult> results = new ArrayList<>();

        for (Map.Entry<Validator.ValidationResource, String> entry : resources.entrySet()) {
            Validator validator = validators.get(entry.getKey());
            if (validator != null) {
                try {
                    if (!validator.validate(entry.getValue(), resources)) {
                        results.add(new ValidationResult(
                                validator.getConfirmStepMessage(entry.getValue(), resources),
                                validator.getFetchFieldName(),
                                validator.getFetchStepMessage(),
                                validator.getErrMessage(entry.getValue()),
                                validator.getFieldType()
                        ));
                    }
                } catch (IllegalStateException cause) {
                    // Access denied (403/WAC token) or other non-recoverable errors are not
                    // validation failures — propagate so the caller can surface the real error.
                    LOGGER.error("Error during validation for resource {}: {}", entry.getKey(), cause.getMessage(), cause);
                    throw cause;
                } catch (Exception cause) {
                    LOGGER.error("Error during validation for resource {}: {}", entry.getKey(), cause.getMessage(), cause);
                    results.add(new ValidationResult("An error occurred during validation. Would you like to try again?",
                            validator.getFetchFieldName(),
                            "Please provide a valid value.",
                            "Validation error: " + cause.getMessage(),
                            validator.getFieldType()
                    ));
                }
            }
        }

        return results;
    }

}
