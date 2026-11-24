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

package app.krista.extensions.essentials.spreadsheets.office365excel.catalog.extresp;

import app.krista.extension.executor.RemediationAction;
import app.krista.extension.executor.impl.AskAPersonAction;
import app.krista.extension.executor.impl.InformAPersonAction;
import app.krista.model.field.NamedField;
import app.krista.model.field.NamedValuedField;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link RemediationActionFactory}.
 */
@DisplayName("RemediationActionFactory")
class RemediationActionFactoryTest {

    @Test
    @DisplayName("createAskAction returns AskAPersonAction with ACTIVE_USER")
    void createAskActionReturnsAskAction() {
        List<NamedField> fields = Collections.emptyList();
        RemediationAction action = RemediationActionFactory.createAskAction("Please provide input", fields);
        assertNotNull(action);
        assertInstanceOf(AskAPersonAction.class, action);
        assertEquals(RemediationAction.RecipientType.ACTIVE_USER, action.getRecipientType());
    }

    @Test
    @DisplayName("createAskAction preserves message")
    void createAskActionPreservesMessage() {
        RemediationAction action = RemediationActionFactory.createAskAction("Test message", List.of());
        assertNotNull(action);
        assertEquals("Test message", action.getMessage());
    }

    @Test
    @DisplayName("createInformAction returns InformAPersonAction with ACTIVE_USER")
    void createInformActionReturnsInformAction() {
        List<NamedValuedField> fields = Collections.emptyList();
        RemediationAction action = RemediationActionFactory.createInformAction("Information message", fields);
        assertNotNull(action);
        assertInstanceOf(InformAPersonAction.class, action);
        assertEquals(RemediationAction.RecipientType.ACTIVE_USER, action.getRecipientType());
    }

    @Test
    @DisplayName("createInformAction preserves message")
    void createInformActionPreservesMessage() {
        RemediationAction action = RemediationActionFactory.createInformAction("Info msg", List.of());
        assertEquals("Info msg", action.getMessage());
    }

    @Test
    @DisplayName("createInformActionALLParticipants returns InformAPersonAction with ALL_PARTICIPANTS")
    void createInformActionAllParticipantsReturnsCorrectRecipient() {
        List<NamedValuedField> fields = Collections.emptyList();
        RemediationAction action = RemediationActionFactory.createInformActionALLParticipants("Broadcast message", fields);
        assertNotNull(action);
        assertInstanceOf(InformAPersonAction.class, action);
        assertEquals(RemediationAction.RecipientType.ALL_PARTICIPANTS, action.getRecipientType());
    }

    @Test
    @DisplayName("createInformActionALLParticipants preserves message")
    void createInformActionAllParticipantsPreservesMessage() {
        RemediationAction action = RemediationActionFactory.createInformActionALLParticipants("Broadcast", List.of());
        assertEquals("Broadcast", action.getMessage());
    }
}
