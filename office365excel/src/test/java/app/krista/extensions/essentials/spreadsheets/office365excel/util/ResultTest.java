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

package app.krista.extensions.essentials.spreadsheets.office365excel.util;

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Drive;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.Identity;
import com.microsoft.graph.models.IdentitySet;
import com.microsoft.graph.models.ItemReference;
import com.microsoft.graph.models.RemoteItem;
import com.microsoft.graph.models.Shared;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.quality.Strictness;
import org.mockito.junit.jupiter.MockitoSettings;

import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class ResultTest {

    @Test
    void getDrives_withValidDriveItems_returnsDriveList() {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = "drive-abc";

        Shared shared = new Shared();
        IdentitySet identitySet = new IdentitySet();
        Identity userIdentity = new Identity();
        userIdentity.displayName = "Alice";
        identitySet.user = userIdentity;
        shared.sharedBy = identitySet;

        remoteItem.parentReference = parentRef;
        remoteItem.shared = shared;
        driveItem.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertEquals(1, drives.size());
        assertEquals("drive-abc", drives.get(0).driveId);
        assertEquals("Alice", drives.get(0).sharedBy);
        assertNull(drives.get(0).createdBy);
    }

    @Test
    void getDrives_withApplicationSharedBy_usesApplicationDisplayName() {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = "drive-app";

        Shared shared = new Shared();
        IdentitySet identitySet = new IdentitySet();
        identitySet.user = null;
        Identity appIdentity = new Identity();
        appIdentity.displayName = "MyApp";
        identitySet.application = appIdentity;
        shared.sharedBy = identitySet;

        remoteItem.parentReference = parentRef;
        remoteItem.shared = shared;
        driveItem.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertEquals(1, drives.size());
        assertEquals("MyApp", drives.get(0).sharedBy);
    }

    @Test
    void getDrives_skipsItemsWithNullRemoteItem() {
        DriveItem driveItemNull = new DriveItem();
        driveItemNull.remoteItem = null;

        DriveItem driveItemValid = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = "drive-valid";
        remoteItem.parentReference = parentRef;
        driveItemValid.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItemNull, driveItemValid));

        assertEquals(1, drives.size());
        assertEquals("drive-valid", drives.get(0).driveId);
    }

    @Test
    void getDrives_skipsItemsWithNullParentReference() {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        remoteItem.parentReference = null;
        driveItem.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertTrue(drives.isEmpty());
    }

    @Test
    void getDrives_skipsDuplicateDriveIds() {
        DriveItem driveItem1 = createDriveItemWithDriveId("drive-dup");
        DriveItem driveItem2 = createDriveItemWithDriveId("drive-dup");
        DriveItem driveItem3 = createDriveItemWithDriveId("drive-unique");

        List<Drive> drives = Result.getDrives(List.of(driveItem1, driveItem2, driveItem3));

        assertEquals(2, drives.size());
        assertEquals("drive-dup", drives.get(0).driveId);
        assertEquals("drive-unique", drives.get(1).driveId);
    }

    @Test
    void getDrives_usesSharePointAsDefaultSharedBy() {
        DriveItem driveItem = createDriveItemWithDriveId("drive-sp");

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertEquals(1, drives.size());
        assertEquals("SharePoint", drives.get(0).sharedBy);
    }

    @Test
    void getDrives_sharedNotNull_sharedByNull_usesSharePoint() {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = "drive-no-sharedby";
        remoteItem.parentReference = parentRef;

        Shared shared = new Shared();
        shared.sharedBy = null;
        remoteItem.shared = shared;
        driveItem.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertEquals(1, drives.size());
        assertEquals("SharePoint", drives.get(0).sharedBy);
    }

    @Test
    void getDrives_sharedByUserAndAppBothNull_usesSharePoint() {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = "drive-noidentity";
        remoteItem.parentReference = parentRef;

        Shared shared = new Shared();
        IdentitySet identitySet = new IdentitySet();
        identitySet.user = null;
        identitySet.application = null;
        shared.sharedBy = identitySet;
        remoteItem.shared = shared;
        driveItem.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertEquals(1, drives.size());
        assertEquals("SharePoint", drives.get(0).sharedBy);
    }

    @Test
    void getDrives_skipsNullDriveId() {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = null;
        remoteItem.parentReference = parentRef;
        driveItem.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertTrue(drives.isEmpty());
    }

    @Test
    void getDrives_skipsBlankDriveId() {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = "   ";
        remoteItem.parentReference = parentRef;
        driveItem.remoteItem = remoteItem;

        List<Drive> drives = Result.getDrives(List.of(driveItem));

        assertTrue(drives.isEmpty());
    }

    @Test
    void getDrives_emptyList_returnsEmptyList() {
        List<Drive> drives = Result.getDrives(Collections.emptyList());

        assertNotNull(drives);
        assertTrue(drives.isEmpty());
    }

    @Test
    void getResultFromMinMaxValueInformation_createsCorrectFreeForm() {
        List<Object> keyColumnData = List.of("Apple", "Banana", "Cherry");
        Map<String, Double> minMaxInfo = new HashMap<>();
        minMaxInfo.put(Constants.MAX_VALUE, 100.0);
        minMaxInfo.put(Constants.MAX_VALUE_ROW_INDEX, 2.0);
        minMaxInfo.put(Constants.MIN_VALUE, 10.0);
        minMaxInfo.put(Constants.MIN_VALUE_ROW_INDEX, 0.0);

        Map<String, Object> result = Result.getResultFromMinMaxValueInformation(keyColumnData, minMaxInfo);

        assertNotNull(result);
        assertEquals(100.0, result.get("Max Value"));
        assertEquals("Cherry", result.get("Max Value Key"));
        assertEquals(10.0, result.get("Min Value"));
        assertEquals("Apple", result.get("Min Value Key"));
    }

    @Test
    void getResultFromMinMaxValueInformation_withNumericKeys_setsNumberType() {
        List<Object> keyColumnData = List.of(1.0, 2.0, 3.0);
        Map<String, Double> minMaxInfo = new HashMap<>();
        minMaxInfo.put(Constants.MAX_VALUE, 300.0);
        minMaxInfo.put(Constants.MAX_VALUE_ROW_INDEX, 2.0);
        minMaxInfo.put(Constants.MIN_VALUE, 100.0);
        minMaxInfo.put(Constants.MIN_VALUE_ROW_INDEX, 0.0);

        Map<String, Object> result = Result.getResultFromMinMaxValueInformation(keyColumnData, minMaxInfo);

        assertNotNull(result);
        assertEquals(300.0, result.get("Max Value"));
        assertEquals(3.0, result.get("Max Value Key"));
        assertEquals(100.0, result.get("Min Value"));
        assertEquals(1.0, result.get("Min Value Key"));
    }

    private DriveItem createDriveItemWithDriveId(String driveId) {
        DriveItem driveItem = new DriveItem();
        RemoteItem remoteItem = new RemoteItem();
        ItemReference parentRef = new ItemReference();
        parentRef.driveId = driveId;
        remoteItem.parentReference = parentRef;
        driveItem.remoteItem = remoteItem;
        return driveItem;
    }
}
