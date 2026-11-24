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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.subcatalog.impl;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ErrorHandlingStateManager;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.errorhandlers.ExtensionResponseGenerator;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.validators.ValidationOrchestrator;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.google.gson.reflect.TypeToken;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.SubCatalogConstants.VALIDATION_RESULTS;

/**
 * Base implementation for reentry handlers with common functionality
 */
public abstract class BaseReentryHandlerImpl {
    protected static final Logger LOGGER = LoggerFactory.getLogger(BaseReentryHandlerImpl.class);

    protected final ErrorHandlingStateManager internalStateManager;
    protected final ExtensionResponseGenerator responseGenerator;

    @Inject
    public BaseReentryHandlerImpl(ErrorHandlingStateManager internalStateManager, ExtensionResponseGenerator responseGenerator) {
        this.internalStateManager = internalStateManager;
        this.responseGenerator = responseGenerator;
    }

    @SuppressWarnings("unchecked")
    protected List<ValidationOrchestrator.ValidationResult> getValidationResults(Map<String, Object> state) {
        if (state == null || !state.containsKey(VALIDATION_RESULTS)) {
            return new ArrayList<>();
        }

        return Constants.GSON.fromJson(Constants.GSON.toJson(state.get(VALIDATION_RESULTS)),
                new TypeToken<List<ValidationOrchestrator.ValidationResult>>() {
                }.getType());
    }
}