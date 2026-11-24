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

import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.Drive;
import app.krista.extensions.essentials.spreadsheets.office365excel.catalog.entities.ExcelFile;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProvider;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.connectors.GraphServiceClientProviderFactory;
import app.krista.extensions.essentials.spreadsheets.office365excel.impl.stores.KristaMediaClient;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.DataModification;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.DriveUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.Result;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.StringUtil;
import app.krista.extensions.essentials.spreadsheets.office365excel.util.constants.Constants;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.DriveItemSearchParameterSet;
import com.microsoft.graph.models.RemoteItem;
import com.microsoft.graph.options.QueryOption;
import com.microsoft.graph.requests.*;
import org.jetbrains.annotations.Nullable;
import org.jvnet.hk2.annotations.Service;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.inject.Inject;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Service
public class DriveService {

    private GraphServiceClientProvider clientProvider;
    private final GraphServiceClientProviderFactory providerFactory;
    private static final Logger LOGGER = LoggerFactory.getLogger(DriveService.class);
    private final KristaMediaClient kristaMediaClient;


    @Inject
    public DriveService(GraphServiceClientProviderFactory providerFactory, KristaMediaClient kristaMediaClient) {
        this.providerFactory = providerFactory;
        this.kristaMediaClient = kristaMediaClient;
    }

    /** Lazy provider creation — avoids MustAuthorizeException during HK2 construction. */
    private GraphServiceClientProvider getClientProvider() {
        if (clientProvider == null) {
            clientProvider = providerFactory.create();
        }
        return clientProvider;
    }

    public List<Drive> getAllSharedDrivesInformation() {
        List<DriveItem> currentDriveItems = getDriveItemsSharedWithMe();
        if (currentDriveItems == null) {
            return Collections.emptyList();
        }
        return Result.getDrives(currentDriveItems);
    }

    public List<ExcelFile> getSharedExcelFilesInformation(String fileName) {
        List<DriveItem> currentDriveItems = getDriveItemsSharedWithMe();
        return currentDriveItems == null ? null : ExcelService.getExcelFiles(fileName, currentDriveItems);
    }

    /**
     * Lists Excel files directly from the given drive using
     * {@code GET /drives/{driveId}/root/children}.
     *
     * <p>This is the correct approach for SharePoint document libraries, which are
     * accessible via site membership but do not appear in {@code /me/drive/sharedWithMe}.</p>
     */
    public List<ExcelFile> getExcelFilesFromDrive(String driveId) {
        try {
            DriveItemCollectionPage page = getClientProvider().getDriveRequestBuilder(driveId)
                    .root().children().buildRequest().top(2000).get();

            List<DriveItem> allItems = new ArrayList<>();
            while (page != null) {
                allItems.addAll(page.getCurrentPage());
                DriveItemCollectionRequestBuilder nextPage = page.getNextPage();
                if (nextPage != null) {
                    page = nextPage.buildRequest().get();
                } else {
                    break;
                }
            }
            LOGGER.info("getExcelFilesFromDrive: {} item(s) found in drive {}", allItems.size(), driveId);
            return ExcelService.getExcelFiles(null, allItems);
        } catch (GraphServiceException cause) {
            LOGGER.error("getExcelFilesFromDrive: failed to list drive {}: {}", driveId, cause.getMessage(), cause);
            return Collections.emptyList();
        }
    }

    /**
     * Searches for Excel files recursively in a drive using the Microsoft Graph search API:
     * {@code GET /drives/{driveId}/root/search(q='{extension}')}.
     *
     * <p>The search API performs server-side recursive traversal across all folders
     * in the drive. Separate searches are made for each Excel file extension
     * (.xlsx, .xls, .csv, .tsv) and results are combined and deduplicated.</p>
     *
     * <p>Pagination follows {@code @odata.nextLink} URLs (containing {@code $skiptoken})
     * returned in each response page. The SDK deserializes {@code @odata.nextLink} into
     * the response's {@code nextLink} field and {@code getNextPage()} returns a request
     * builder targeting that URL.</p>
     */
    public List<ExcelFile> searchExcelFilesInDrive(String driveId, String fileFormat) {
        String[] extensions = (fileFormat != null && !fileFormat.isBlank())
                ? new String[]{fileFormat}
                : new String[]{".xlsx", ".xls", ".csv", ".tsv"};
        LinkedHashMap<String, DriveItem> uniqueItems = new LinkedHashMap<>();

        // Acquire the Graph client once and reuse for all extension searches + pagination
        GraphServiceClient<?> graphClient;
        try {
            graphClient = getClientProvider().getGraphServiceClient();
        } catch (IOException e) {
            LOGGER.error("searchExcelFilesInDrive: failed to get Graph client: {}", e.getMessage(), e);
            return Collections.emptyList();
        }

        for (String ext : extensions) {
            try {
                DriveItemSearchParameterSet params = DriveItemSearchParameterSet.newBuilder().withQ(ext).build();
                DriveItemSearchCollectionPage page = graphClient.me().drives(driveId)
                        .root().search(params).buildRequest().top(200).get();
                while (page != null) {
                    for (DriveItem item : page.getCurrentPage()) {
                        if (item.id != null) {
                            uniqueItems.putIfAbsent(item.id, item);
                        }
                    }
                    // SDK follows @odata.nextLink ($skiptoken) via getNextPage()
                    DriveItemSearchCollectionRequestBuilder nextPage = page.getNextPage();
                    if (nextPage != null) {
                        page = nextPage.buildRequest().get();
                    } else {
                        break;
                    }
                }
                LOGGER.info("searchExcelFilesInDrive: completed search for '{}' in drive {}", ext, driveId);
            } catch (GraphServiceException cause) {
                LOGGER.error("searchExcelFilesInDrive: search for '{}' failed in drive {}: {}",
                        ext, driveId, cause.getMessage(), cause);
            }
        }

        LOGGER.info("searchExcelFilesInDrive: {} unique item(s) found in drive {}", uniqueItems.size(), driveId);
        return ExcelService.getExcelFiles(null, new ArrayList<>(uniqueItems.values()));
    }

    /**
     * Fetches metadata for a specific drive by its ID using
     * {@code GET /drives/{driveId}} and returns it as a single-element list.
     *
     * <p>Used when the caller already knows the drive ID (e.g. a SharePoint library
     * that does not appear in {@code /me/drive/sharedWithMe}).</p>
     */
    public List<Drive> getDriveById(String driveId) {
        if (driveId == null || driveId.isBlank()) {
            return Collections.emptyList();
        }
        try {
            com.microsoft.graph.models.Drive graphDrive = getClientProvider().getDriveRequestBuilder(driveId)
                    .buildRequest().get();
            if (graphDrive == null) {
                LOGGER.info("getDriveById: no response for driveId={}", driveId);
                return Collections.emptyList();
            }
            String name = graphDrive.name != null ? graphDrive.name : "SharePoint Drive";
            LOGGER.info("getDriveById: found drive {} name={}", driveId, name);
            return List.of(new Drive(driveId, name, null));
        } catch (GraphServiceException cause) {
            LOGGER.error("getDriveById: failed for driveId={}: {}", driveId, cause.getMessage(), cause);
            return Collections.emptyList();
        }
    }

    private List<DriveItem> getDriveItemsSharedWithMe() {
        try {
            DriveSharedWithMeCollectionPage sharedWithMeCollectionPage = getClientProvider().getDriveRequestBuilder(null).sharedWithMe()
                    .buildRequest()
                    .top(2000)
                    .get();

            List<DriveItem> currentDriveItems = new ArrayList<>();
            while (sharedWithMeCollectionPage != null) {
                currentDriveItems.addAll(sharedWithMeCollectionPage.getCurrentPage());
                DriveSharedWithMeCollectionRequestBuilder nextPage = sharedWithMeCollectionPage.getNextPage();
                if (nextPage != null) {
                    sharedWithMeCollectionPage = nextPage.buildRequest().get();
                } else { break; }
            }
            long excelCount  = currentDriveItems.stream().filter(i -> i.file   != null && i.name != null && StringUtil.isExcelFile(i.name)).count();
            long folderCount = currentDriveItems.stream().filter(i -> i.folder != null).count();
            long otherCount  = currentDriveItems.size() - excelCount - folderCount;
            LOGGER.info("getDriveItemsSharedWithMe: total={} (Excel files={}, folders={}, other={})",
                    currentDriveItems.size(), excelCount, folderCount, otherCount);
            for (int idx = 0; idx < currentDriveItems.size(); idx++) {
                DriveItem item = currentDriveItems.get(idx);
                String type = item.folder != null ? "FOLDER" : (item.file != null ? "FILE" : "OTHER");
                String remoteInfo = item.remoteItem != null && item.remoteItem.parentReference != null
                        ? "remoteItem.driveId=" + item.remoteItem.parentReference.driveId
                        : (item.remoteItem != null ? "remoteItem(no parentRef)" : "no-remoteItem");
                LOGGER.info("getDriveItemsSharedWithMe: [{}] name='{}' type={} {}", idx + 1, item.name, type, remoteInfo);
            }
            return currentDriveItems;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to get drive item: {}", cause.getMessage(), cause);
            return null;
        }
    }

    public List<ExcelFile> getExcelFileFromSharedDrive(String sharedFolderName, String sharedFileName) {

        if (sharedFolderName == null || sharedFolderName.isBlank()) {
            List<ExcelFile> sharedExcelFilesInformation = getSharedExcelFilesInformation(sharedFileName);
            if (sharedExcelFilesInformation != null && sharedExcelFilesInformation.isEmpty()) {
                LOGGER.info("Shared file {} is not accessible.", sharedFileName);
            }
            return sharedExcelFilesInformation;
        }

        String[] folderPath = sharedFolderName.split(Constants.BACKWARD_SLASH);
        List<DriveItem> driveItemsSharedWithMe = getDriveItemsSharedWithMe();
        if (driveItemsSharedWithMe == null) {
            return Collections.emptyList();
        }

        DriveItem firstDriveItemFolder = null;
        // Check folder is available at first level
        for (DriveItem driveItem : driveItemsSharedWithMe) {
            if ((driveItem.folder != null && driveItem.name != null && driveItem.name.equalsIgnoreCase(folderPath[0]))) {
                // As we get folder name, no need to check other DriveItem information
                // TODO: Need to cover, multiple folder "Demo" shared by multiple people

                // Folder "Demo" shared by single person
                firstDriveItemFolder = driveItem;
                break;
            }
        }

        if (firstDriveItemFolder == null) {
            LOGGER.info("Shared folder {} is not accessible", folderPath[0]);
            return Collections.emptyList();
        }

        if (folderPath.length == 1) {
            return getFileFromFirstLevelSharedFolder(sharedFileName, firstDriveItemFolder);
        }
        return getFileFromNestedSharedFolder(sharedFileName, folderPath, firstDriveItemFolder);

    }

    private List<ExcelFile> getFileFromNestedSharedFolder(String sharedFileName, String[] folderPath, DriveItem firstDriveItemFolder) {
        if (firstDriveItemFolder == null || firstDriveItemFolder.remoteItem == null || firstDriveItemFolder.remoteItem.parentReference == null) {
            return null;
        }

        List<ExcelFile> excelFiles = new ArrayList<>();
        boolean folderAvailable = false;
        DriveItem currentFolder = firstDriveItemFolder;
        String driveId = firstDriveItemFolder.remoteItem.parentReference.driveId;
        List<DriveItem> folderItems = null;

        for (int i = 1; i < folderPath.length + 1; i++) {
            DriveRequestBuilder driveRequestBuilder1 = getClientProvider().getDriveRequestBuilder(driveId);
            folderItems = getFolderItems(driveRequestBuilder1, currentFolder.id);
            if (folderItems != null) {
                for (DriveItem driveItem : folderItems) {
                    if ((driveItem.folder != null && driveItem.name != null && driveItem.name.equalsIgnoreCase(folderPath[i]))) {
                        folderAvailable = true;
                        currentFolder = driveItem;
                        break;
                    }
                }
            }
            if (!folderAvailable) {
                LOGGER.info("Shared folder {} is not accessible.", folderPath[i]);
                break;
            }
        }

        if (folderAvailable) {
            excelFiles = getExcelFileFromSharedFolder(sharedFileName, folderItems);
            if (excelFiles == null) {
                throw new IllegalArgumentException("Shared file is not accessible");
            }
            if (excelFiles.isEmpty()) {
                LOGGER.info("Shared file {} is not accessible.", sharedFileName);
            }
        }
        return excelFiles;
    }

    @Nullable
    private List<ExcelFile> getFileFromFirstLevelSharedFolder(String sharedFileName, DriveItem firstDriveItemFolder) {
        List<ExcelFile> excelFiles = new ArrayList<>();
        // If only one folder name is given and that is available at first level
        if (firstDriveItemFolder != null) {
            RemoteItem remoteItem = firstDriveItemFolder.remoteItem;
            if (remoteItem != null && remoteItem.parentReference != null) {
                DriveRequestBuilder driveRequestBuilder1 = getClientProvider().getDriveRequestBuilder(remoteItem.parentReference.driveId);
                List<DriveItem> folderItems = getFolderItems(driveRequestBuilder1, firstDriveItemFolder.id);
                excelFiles = getExcelFileFromSharedFolder(sharedFileName, folderItems);
            }
        } else {
            // If any folder name is not available, then path is invalid
            return null;
        }
        if (excelFiles == null) {
            throw new IllegalArgumentException("Shared File is not accessible");
        }
        if (excelFiles.isEmpty()) {
            LOGGER.info("Shared file {} is not accessible.", sharedFileName);
        }
        return excelFiles;
    }

    private List<DriveItem> getFolderItems(DriveRequestBuilder driveRequestBuilder, String folderId) {
        try {
            DriveItemCollectionPage children = driveRequestBuilder.items(folderId).children().buildRequest().get();

            List<DriveItem> currentDriveItems = new ArrayList<>();
            while (children != null) {
                currentDriveItems.addAll(children.getCurrentPage());
                DriveItemCollectionRequestBuilder nextPage = children.getNextPage();
                if (nextPage != null) {
                    children = nextPage.buildRequest().get();
                } else {
                    break;
                }
            }

            return currentDriveItems;
        } catch (GraphServiceException cause) {
            LOGGER.error("Failed to get folder items: {}", cause.getMessage(), cause);
            return null;
        }
    }

    private List<ExcelFile> getExcelFileFromSharedFolder(String fileName, List<DriveItem> currentDriveItems) {
        if (StringUtil.isNullOrEmpty(fileName) && currentDriveItems == null) {
            return null;
        }

        List<ExcelFile> excelFiles = new ArrayList<>();

        for (DriveItem driveItem : currentDriveItems) {
            if (driveItem.name != null && driveItem.name.equalsIgnoreCase(fileName)) {
                String fileDriveId = driveItem.parentReference == null ? null :
                        driveItem.parentReference.driveId == null ? null : driveItem.parentReference.driveId;
                excelFiles.add(new ExcelFile(driveItem.id, fileDriveId, driveItem.name, driveItem.webUrl));
            }
        }
        return excelFiles;
    }

    public ExcelFile getExcelFileByPath(String sourcePath, String fileSearch) {
        if (sourcePath == null || sourcePath.isBlank()) {
            throw new IllegalArgumentException("Source path must not be blank");
        }

        // URL-decode (handles %20 etc.)
        String decoded;
        try {
            decoded = java.net.URLDecoder.decode(sourcePath.trim(), java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception e) {
            decoded = sourcePath.trim();
        }

        // Normalise leading slash
        if (decoded.startsWith("/")) {
            decoded = decoded.substring(1);
        }

        // Split into folder path and filename
        int lastSlash = decoded.lastIndexOf('/');
        String folderPath = lastSlash > 0 ? decoded.substring(0, lastSlash) : "";
        String fileName   = lastSlash >= 0 ? decoded.substring(lastSlash + 1) : decoded;

        if (fileName.isBlank()) {
            throw new IllegalArgumentException("Source path must include a filename");
        }

        if (fileSearch == null || fileSearch.isBlank()) {
            throw new IllegalArgumentException(
                    "File Search is required. Choose one of: My Drive, Site Files");
        }

        LOGGER.info("getExcelFileByPath: folderPath='{}', fileName='{}', fileSearch='{}'",
                folderPath, fileName, fileSearch);

        // My Drive: resolve by path directly in the user's personal OneDrive
        if ("My Drive".equalsIgnoreCase(fileSearch)) {
            String itemPath = folderPath.isBlank() ? fileName : folderPath + "/" + fileName;
            LOGGER.info("getExcelFileByPath: My Drive lookup: itemPath='{}'", itemPath);
            ExcelFile file = getExcelFileFromMyDrive(itemPath);
            if (file != null) {
                return file;
            }
            throw new IllegalArgumentException("File not found in My Drive at path: " + sourcePath);
        }

        // Site Files: path must be site/library[/subfolders]/filename
        if ("Site Files".equalsIgnoreCase(fileSearch)) {
            String[] segments = folderPath.split("/");
            if (segments.length < 2) {
                throw new IllegalArgumentException(
                        "Site Files lookup requires at least two path segments (site/library): " + sourcePath);
            }
            String siteName    = segments[0];
            String libraryName = segments[1];
            String itemPath    = (segments.length > 2
                    ? String.join("/", Arrays.copyOfRange(segments, 2, segments.length)) + "/"
                    : "") + fileName;
            LOGGER.info("getExcelFileByPath: site='{}', library='{}', itemPath='{}'",
                    siteName, libraryName, itemPath);
            ExcelFile siteFile = getExcelFileFromSiteByPath(siteName, libraryName, itemPath);
            if (siteFile != null) {
                return siteFile;
            }
            throw new IllegalArgumentException(
                    "File not found in Site Files at path: " + sourcePath);
        }

        throw new IllegalArgumentException(
                "Unknown File Search value '" + fileSearch + "'. Choose one of: My Drive, Site Files");
    }

    /**
     * Resolves a file from the signed-in user's personal OneDrive using
     * {@code GET /me/drive/root:/{itemPath}}.
     *
     * <p>Use this when the file is owned by the authenticated user, not shared
     * by someone else and not inside a SharePoint site library.</p>
     */
    @Nullable
    private ExcelFile getExcelFileFromMyDrive(String itemPath) {
        try {
            GraphServiceClient<?> graphClient = getClientProvider().getGraphServiceClient();
            DriveItem item = graphClient.me().drive().root().itemWithPath(itemPath).buildRequest().get();
            if (item == null) {
                LOGGER.info("getExcelFileFromMyDrive: item not found at '{}'", itemPath);
                return null;
            }
            String driveId = item.parentReference != null ? item.parentReference.driveId : null;
            LOGGER.info("getExcelFileFromMyDrive: found item '{}' id={} driveId={}", item.name, item.id, driveId);
            return new ExcelFile(item.id, driveId, item.name, item.webUrl);
        } catch (GraphServiceException e) {
            LOGGER.error("getExcelFileFromMyDrive: Graph API error - {}", e.getMessage(), e);
            return null;
        } catch (IOException e) {
            LOGGER.error("getExcelFileFromMyDrive: failed to get Graph client - {}", e.getMessage(), e);
            return null;
        }
    }

    /**
     * Resolves a SharePoint file via the Sites API:
     * 1. Searches for a site whose display name matches {@code siteName}.
     * 2. Finds the document library (drive) whose name matches {@code libraryName}.
     * 3. Uses {@code GET /drives/{driveId}/root:/{itemPath}} to fetch the file metadata.
     *
     * <p>This handles SharePoint paths where the first segment is a site name rather than
     * a folder that appears in {@code /me/drive/sharedWithMe}.</p>
     */
    @Nullable
    private ExcelFile getExcelFileFromSiteByPath(String siteName, String libraryName, String itemPath) {
        GraphServiceClient<?> graphClient;
        try {
            graphClient = getClientProvider().getGraphServiceClient();
        } catch (IOException e) {
            LOGGER.error("getExcelFileFromSiteByPath: failed to get Graph client: {}", e.getMessage(), e);
            return null;
        }

        try {
            String siteId = resolveSiteId(graphClient, siteName);
            if (siteId == null) return null;

            String driveId = resolveDriveId(graphClient, siteId, libraryName);
            if (driveId == null) return null;

            return resolveFileInDrive(graphClient, driveId, itemPath);
        } catch (GraphServiceException e) {
            LOGGER.error("getExcelFileFromSiteByPath: Graph API error - {}", e.getMessage(), e);
            return null;
        }
    }

    /**
     * Step 1: Search for a SharePoint site by display name and return its Graph API site ID.
     * Prefers an exact case-insensitive match on {@code displayName}; falls back to the first result.
     *
     * @return the site ID, or {@code null} if no site is found
     */
    @Nullable
    private String resolveSiteId(GraphServiceClient<?> graphClient, String siteName) {
        List<QueryOption> searchOption = Collections.singletonList(new QueryOption("search", siteName));
        SiteCollectionPage sitePage = graphClient.sites().buildRequest(searchOption).get();
        if (sitePage == null || sitePage.getCurrentPage().isEmpty()) {
            LOGGER.info("getExcelFileFromSiteByPath: no site found matching '{}'", siteName);
            return null;
        }

        com.microsoft.graph.models.Site matchedSite = null;
        for (com.microsoft.graph.models.Site site : sitePage.getCurrentPage()) {
            if (site.displayName != null && site.displayName.equalsIgnoreCase(siteName)) {
                matchedSite = site;
                break;
            }
        }
        if (matchedSite == null) {
            matchedSite = sitePage.getCurrentPage().get(0);
        }
        String siteId = matchedSite.id;
        LOGGER.info("getExcelFileFromSiteByPath: resolved site '{}' -> siteId={}", siteName, siteId);
        return siteId;
    }

    /**
     * Step 2: Find the document library (drive) inside a SharePoint site that matches the given
     * library name. Matches on {@code drive.name} (display name) OR the last URL segment of
     * {@code drive.webUrl} (the URL-accessible name, e.g. "Shared Documents"), because they often differ.
     *
     * @return the drive ID, or {@code null} if no matching drive is found
     */
    @Nullable
    private String resolveDriveId(GraphServiceClient<?> graphClient, String siteId, String libraryName) {
        DriveCollectionPage drivesPage = graphClient.sites(siteId).drives().buildRequest().get();
        if (drivesPage == null) {
            LOGGER.info("getExcelFileFromSiteByPath: no drives found for site '{}'", siteId);
            return null;
        }

        com.microsoft.graph.models.Drive matchedDrive = null;
        for (com.microsoft.graph.models.Drive drive : drivesPage.getCurrentPage()) {
            String urlSegment = extractLastUrlSegment(drive.webUrl);
            LOGGER.info("getExcelFileFromSiteByPath: available drive name='{}', webUrlSegment='{}'",
                    drive.name, urlSegment);
            boolean nameMatch = drive.name != null && drive.name.equalsIgnoreCase(libraryName);
            boolean urlMatch  = urlSegment != null && urlSegment.equalsIgnoreCase(libraryName);
            if (nameMatch || urlMatch) {
                matchedDrive = drive;
                break;
            }
        }
        if (matchedDrive == null) {
            LOGGER.info("getExcelFileFromSiteByPath: no drive matching library '{}' in site '{}' " +
                    "(checked both drive.name and webUrl segment)", libraryName, siteId);
            return null;
        }
        String driveId = matchedDrive.id;
        LOGGER.info("getExcelFileFromSiteByPath: resolved library '{}' -> driveId={}", libraryName, driveId);
        return driveId;
    }

    /**
     * Step 3: Resolve a file by its path within a drive and return it as an {@link ExcelFile}.
     *
     * @return the {@link ExcelFile}, or {@code null} if no item exists at the given path
     */
    @Nullable
    private ExcelFile resolveFileInDrive(GraphServiceClient<?> graphClient, String driveId, String itemPath) {
        DriveItem item = graphClient.drives(driveId).root().itemWithPath(itemPath).buildRequest().get();
        if (item == null) {
            LOGGER.info("getExcelFileFromSiteByPath: item not found at '{}' in drive '{}'", itemPath, driveId);
            return null;
        }
        LOGGER.info("getExcelFileFromSiteByPath: found item '{}' id={}", item.name, item.id);
        return new ExcelFile(item.id, driveId, item.name, item.webUrl);
    }

    /**
     * Returns the last path segment of a URL after URL-decoding it.
     * e.g. "https://tenant.sharepoint.com/sites/Eng/Shared%20Documents" → "Shared Documents"
     */
    @Nullable
    private static String extractLastUrlSegment(String webUrl) {
        if (webUrl == null || webUrl.isBlank()) return null;
        try {
            String decoded = java.net.URLDecoder.decode(webUrl, java.nio.charset.StandardCharsets.UTF_8);
            int slash = decoded.lastIndexOf('/');
            return slash >= 0 ? decoded.substring(slash + 1) : decoded;
        } catch (Exception e) {
            return null;
        }
    }

    public app.krista.model.base.File getPdfFile(String fileId, String driveId) throws IOException {
        if (fileId == null) {
            LOGGER.info("Failed to get Pdf file. Given Id is invalid.");
            return null;
        }

        try {

            InputStream inputStream = getInputStreamForExcelFile(fileId, driveId);

            String fileName = "doc_pdf_" + UUID.randomUUID() + ".pdf";
            java.io.File pdfFile = DriveUtil.copyInputStreamToFile(inputStream, fileName);
            if (pdfFile == null) {
                throw new IllegalArgumentException("Failed to create PDF file.");
            }
            return kristaMediaClient.toKristaFile(pdfFile);
        } catch (IOException | GraphServiceException cause) {
            LOGGER.info("Failed to get PDF file {}", cause.getMessage(), cause);
            return null;
        }

    }

    @Nullable
    private InputStream getInputStreamForExcelFile(String fileId, String driveId) {
        LinkedList<QueryOption> requestOptions = new LinkedList<>();
        requestOptions.add(new QueryOption("format", "pdf"));

        return getClientProvider().getDriveRequestBuilder(driveId).items(fileId).content().buildRequest(requestOptions).get();
    }

    public Map<String, Object> getPdfFileForGivenExcel(app.krista.model.base.File excelFile, String path) {


        if (excelFile == null || excelFile.getExtension() == null || !excelFile.getExtension().equalsIgnoreCase("xlsx")) {
            LOGGER.info("Failed to get PDF File. Upload Excel File.");
            return Collections.emptyMap();
        }

        try {
            String pdfFileId = uploadFileToDrive(kristaMediaClient, excelFile, path);
            app.krista.model.base.File pdfFile = getPdfFile(pdfFileId, null);

            return DataModification.getResultMap(pdfFileId, pdfFile);
        } catch (IOException | GraphServiceException cause) {
            LOGGER.info("Failed to get PDF file {}", cause.getMessage(), cause);
            return Map.of();
        }
    }

    private String uploadFileToDrive(KristaMediaClient kristaMediaClient, app.krista.model.base.File file, String path) throws IOException {
        java.io.File parsedFile = kristaMediaClient.toJavaFile(file);
        String finalPath = path == null ? parsedFile.getName() : path + "/" + parsedFile.getName();
        final DriveItem driveItem = getClientProvider().getDriveRequestBuilder(null).root().itemWithPath(finalPath).content()
                .buildRequest()
                .put(DriveUtil.getByteArray(parsedFile));
        return driveItem != null ? driveItem.id : null;
    }

}
