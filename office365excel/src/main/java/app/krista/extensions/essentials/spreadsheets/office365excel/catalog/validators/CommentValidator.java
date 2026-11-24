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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.FieldTypes;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

import static app.krista.extensions.essentials.spreadsheets.office365excel.catalog.util.Office365ExcelResources.COMMENT;

/**
 * Validator for comments in Office365Excel extension.
 */
public class CommentValidator implements Validator {

    private static final Logger LOGGER = LoggerFactory.getLogger(CommentValidator.class);
    private static final int MAX_COMMENT_LENGTH = 1024; // Reasonable limit for comments

    @Override
    public Boolean validate(String comment, Map<ValidationResource, String> context) {
        if (comment == null) {
            LOGGER.debug("Comment validation failed: comment is null");
            return false;
        }
        
        try {
            // Check length constraint
            if (comment.length() > MAX_COMMENT_LENGTH) {
                LOGGER.debug("Comment validation failed: comment exceeds maximum length of {} characters", MAX_COMMENT_LENGTH);
                return false;
            }
            
            return true;
        } catch (Exception cause) {
            LOGGER.error("Error validating comment: {}", cause.getMessage(), cause);
            return false;
        }
    }

    @Override
    public String getFetchFieldName() {
        return COMMENT;
    }

    @Override
    public String getFieldType() {
        return FieldTypes.TEXT_FIELD_LOGICAL_NAME;
    }

    @Override
    public String getFetchStepMessage() {
        return "Please enter a comment (optional).";
    }

    @Override
    public String getConfirmStepMessage(String comment, Map<ValidationResource, String> context) {
        return "The comment is too long. Would you like to enter a shorter comment?";
    }

    @Override
    public String getErrMessage(String comment) {
        return "Comment is too long. Please limit your comment to " + MAX_COMMENT_LENGTH + " characters.";
    }
}