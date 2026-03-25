import React, { useState, useEffect, useCallback } from "react";
import {
  Button,
  Dropdown,
  Option,
  Checkbox,
  Spinner,
  Text,
  Input,
  Divider,
  makeStyles,
  tokens,
  shorthands,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  DismissCircle24Filled,
  Search20Regular,
} from "@fluentui/react-icons";

import type {
  SharePointSite,
  Library,
  ProjectFolder,
  SubFolder,
  MailFolder,
  SaveStatus,
  ConversationHistory,
  FileNameField,
} from "../types";
import { DEFAULT_FILENAME_CONFIG } from "../types";
import { APP_CONFIG } from "../config";
import { FileNameBuilder } from "./FileNameBuilder";
import {
  searchSites,
  getLibraries,
  getRootFolders,
  getProjectFolders,
  getSubFolders,
  getMailMimeContent,
  uploadToSharePoint,
  getMailFolders,
  moveMailToFolder,
  generateEmailFileName,
} from "../services/graph";
import {
  getHistoryForConversation,
  saveConversationHistory,
  loadConversationHistory,
  loadFileNameConfig,
  saveFileNameConfig,
} from "../services/roamingSettings";

// ============================================================
// Styles
// ============================================================

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    ...shorthands.padding("16px"),
    boxSizing: "border-box",
    fontFamily: tokens.fontFamilyBase,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    display: "flex",
    alignItems: "center",
    marginBottom: "16px",
    ...shorthands.gap("8px"),
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  form: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("12px"),
    flexGrow: 1,
    overflowY: "auto" as const,
  },
  fieldGroup: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("4px"),
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    textTransform: "uppercase" as const,
    letterSpacing: "0.5px",
  },
  searchRow: {
    display: "flex",
    ...shorthands.gap("4px"),
  },
  footer: {
    marginTop: "auto",
    paddingTop: "12px",
  },
  saveButton: {
    width: "100%",
    height: "40px",
  },
  statusContainer: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    ...shorthands.gap("8px"),
    ...shorthands.padding("12px"),
    ...shorthands.borderRadius("8px"),
    marginTop: "8px",
  },
  statusSuccess: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  statusError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  prefillBadge: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorPaletteBlueForeground2,
    fontStyle: "italic" as const,
  },
});

// ============================================================
// Helpers
// ============================================================

function normalizeSubject(subject: string): string {
  return subject
    .replace(/^(re|fw|fwd|tr|antw):\s*/gi, "")
    .trim()
    .toLowerCase();
}

// ============================================================
// Component
// ============================================================

export const Taskpane: React.FC = () => {
  const styles = useStyles();

  // Site
  const [sites, setSites] = useState<SharePointSite[]>([]);
  const [selectedSiteId, setSelectedSiteId] = useState<string>("");
  const [siteSearch, setSiteSearch] = useState<string>("");
  const [loadingSites, setLoadingSites] = useState(false);

  // Bibliotheek
  const [libraries, setLibraries] = useState<Library[]>([]);
  const [selectedLibraryId, setSelectedLibraryId] = useState<string>("");
  const [loadingLibraries, setLoadingLibraries] = useState(false);

  // Submap binnen bibliotheek
  const [rootFolders, setRootFolders] = useState<SubFolder[]>([]);
  const [selectedRootFolderId, setSelectedRootFolderId] = useState<string>("");
  const [loadingRootFolders, setLoadingRootFolders] = useState(false);

  // Project
  const [projects, setProjects] = useState<ProjectFolder[]>([]);
  const [selectedProjectId, setSelectedProjectId] = useState<string>("");
  const [projectSearch, setProjectSearch] = useState<string>("");
  const [loadingProjects, setLoadingProjects] = useState(false);

  // Submap van project
  const [subFolders, setSubFolders] = useState<SubFolder[]>([]);
  const [selectedSubFolderId, setSelectedSubFolderId] = useState<string>("");
  const [loadingSubFolders, setLoadingSubFolders] = useState(false);

  // Mail folders
  const [mailFolders, setMailFolders] = useState<MailFolder[]>([]);
  const [selectedArchiveFolderId, setSelectedArchiveFolderId] = useState<string>("");

  // Bestandsnaam
  const [fileNameFields, setFileNameFields] = useState<FileNameField[]>(
    () => loadFileNameConfig().fields
  );
  const [fileName, setFileName] = useState<string>("");
  const [mailSubject, setMailSubject] = useState<string>("");
  const [mailDate, setMailDate] = useState<Date>(new Date());
  const [mailSender, setMailSender] = useState<string>("");
  const [mailRecipient, setMailRecipient] = useState<string>("");

  // Status
  const [saveAttachments, setSaveAttachments] = useState<boolean>(true);
  const [saveStatus, setSaveStatus] = useState<SaveStatus>("idle");
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [isPrefilled, setIsPrefilled] = useState<boolean>(false);

  // --- Helpers ---
  const getMailItem = () => Office.context.mailbox.item;
  const getConversationId = (): string | undefined => getMailItem()?.conversationId;

  // --- Initialisatie ---
  useEffect(() => {
    initializeTaskpane();
  }, []);

  const initializeTaskpane = async () => {
    try {
      const item = getMailItem();
      if (item) {
        const subject = item.subject || "Geen onderwerp";
        const dateReceived = item.dateTimeCreated
          ? new Date(item.dateTimeCreated.toString())
          : new Date();
        const sender = item.from?.emailAddress || "";
        const recipient = (item as any).to?.[0]?.emailAddress || "";
        const fields = loadFileNameConfig().fields;

        setMailSubject(subject);
        setMailDate(dateReceived);
        setMailSender(sender);
        setMailRecipient(recipient);
        setFileName(generateEmailFileName(subject, dateReceived, fields, sender, recipient));
      }

      await loadAllSites();
      loadMailFolders();
    } catch (error) {
      console.error("Initialisatie fout:", error);
    }
  };

  // --- Sites laden ---
  const loadAllSites = async () => {
    setLoadingSites(true);
    try {
      const results = await searchSites("*");
      setSites(results);

      // Prefill: eerst conversationId, dan genormaliseerd onderwerp
      const conversationId = getConversationId();
      let history: ConversationHistory | undefined;

      if (conversationId) {
        history = getHistoryForConversation(conversationId);
      }
      if (!history) {
        const subject = getMailItem()?.subject || "";
        const normalized = normalizeSubject(subject);
        if (normalized) {
          history = loadConversationHistory().find((h) => h.normalizedSubject === normalized);
        }
      }

      if (history) {
        await applyHistory(history, results);
      } else {
        const defaultSite = results.find((s) =>
          s.webUrl?.toLowerCase().includes(APP_CONFIG.defaultSiteHostname.toLowerCase())
        );
        if (defaultSite) {
          setSelectedSiteId(defaultSite.id);
          await loadLibrariesForSite(defaultSite.id);
        }
      }
    } catch (error) {
      console.error("Kan sites niet laden:", error);
    } finally {
      setLoadingSites(false);
    }
  };

  const loadLibrariesForSite = async (siteId: string): Promise<Library[]> => {
    setLoadingLibraries(true);
    try {
      const libs = await getLibraries(siteId);
      setLibraries(libs);
      return libs;
    } catch (error) {
      console.error("Kan bibliotheken niet laden:", error);
      setLibraries([]);
      return [];
    } finally {
      setLoadingLibraries(false);
    }
  };

  const loadRootFoldersForLibrary = async (siteId: string, libraryId: string): Promise<SubFolder[]> => {
    setLoadingRootFolders(true);
    try {
      const folders = await getRootFolders(siteId, libraryId);
      setRootFolders(folders);
      return folders;
    } catch (error) {
      console.error("Kan rootmappen niet laden:", error);
      setRootFolders([]);
      return [];
    } finally {
      setLoadingRootFolders(false);
    }
  };

  const loadProjects = async (siteId: string, libraryId: string, subPath?: string) => {
    setLoadingProjects(true);
    try {
      console.log("[loadProjects]", { siteId, libraryId, subPath });
      const folders = await getProjectFolders(siteId, libraryId, subPath);
      console.log("[loadProjects] resultaat:", folders.length, "mappen");
      setProjects(folders);
    } catch (error: any) {
      console.error("[loadProjects] FOUT:", error?.message || error);
      setErrorMessage(`Projecten laden mislukt: ${error?.message || error}`);
      setSaveStatus("error");
      setProjects([]);
    } finally {
      setLoadingProjects(false);
    }
  };

  const loadProjectSubFolders = async (siteId: string, projectId: string) => {
    setLoadingSubFolders(true);
    try {
      const folders = await getSubFolders(siteId, projectId);
      setSubFolders(folders);
    } catch (error) {
      console.error("Kan submappen niet laden:", error);
      setSubFolders([]);
    } finally {
      setLoadingSubFolders(false);
    }
  };

  const loadMailFolders = async () => {
    try {
      setMailFolders(await getMailFolders());
    } catch (error) {
      console.error("Kan mail folders niet laden:", error);
    }
  };

  // --- Bestandsnaam ---
  const handleFileNameFieldsChange = async (fields: FileNameField[]) => {
    setFileNameFields(fields);
    setFileName(generateEmailFileName(mailSubject, mailDate, fields, mailSender, mailRecipient));
    await saveFileNameConfig({ fields }).catch(console.error);
  };

  // --- Smart Prefill ---
  const applyHistory = async (history: ConversationHistory, availableSites: SharePointSite[]) => {
    if (!availableSites.find((s) => s.id === history.siteId)) return;
    setIsPrefilled(true);
    setSelectedSiteId(history.siteId);

    const libs = await loadLibrariesForSite(history.siteId);
    if (!libs.find((l) => l.id === history.libraryId)) return;
    setSelectedLibraryId(history.libraryId);

    const rootFolderList = await loadRootFoldersForLibrary(history.siteId, history.libraryId);

    // Probeer submap te matchen op naam
    const matchedRoot = rootFolderList.find((f) =>
      history.projectFolderName?.startsWith(f.name) === false &&
      history.libraryName === f.name
    );
    const subFolderId = matchedRoot?.id || "";
    setSelectedRootFolderId(subFolderId);

    await loadProjects(history.siteId, history.libraryId, subFolderId || undefined);
    setSelectedProjectId(history.projectFolderId);
    await loadProjectSubFolders(history.siteId, history.projectFolderId);
    setSelectedSubFolderId(history.subFolderId);

    if (history.archiveMailFolderId) setSelectedArchiveFolderId(history.archiveMailFolderId);
  };

  // --- Event handlers ---
  const handleSiteSearch = async () => {
    if (!siteSearch.trim()) return;
    setLoadingSites(true);
    try {
      setSites(await searchSites(siteSearch));
    } catch (error) {
      console.error("Site zoeken mislukt:", error);
    } finally {
      setLoadingSites(false);
    }
  };

  const handleSiteChange = async (_: any, data: any) => {
    const siteId = data.optionValue;
    setSelectedSiteId(siteId);
    setSelectedLibraryId("");
    setRootFolders([]);
    setSelectedRootFolderId("");
    setProjects([]);
    setSelectedProjectId("");
    setSubFolders([]);
    setSelectedSubFolderId("");
    setIsPrefilled(false);
    await loadLibrariesForSite(siteId);
  };

  const handleLibraryChange = async (_: any, data: any) => {
    const libraryId = data.optionValue;
    setSelectedLibraryId(libraryId);
    setRootFolders([]);
    setSelectedRootFolderId("");
    setProjects([]);
    setSelectedProjectId("");
    setSubFolders([]);
    setSelectedSubFolderId("");
    setIsPrefilled(false);

    // Eerst rootmappen laden, dan projecten laden (met eventuele auto-selectie)
    const rootFolderList = await loadRootFoldersForLibrary(selectedSiteId, libraryId);

    if (rootFolderList.length === 1) {
      // Automatisch de enige submap selecteren en projecten daarbinnen laden
      setSelectedRootFolderId(rootFolderList[0].id);
      await loadProjects(selectedSiteId, libraryId, rootFolderList[0].name);
    } else {
      // Meerdere of geen submappen: laad projecten vanuit root
      await loadProjects(selectedSiteId, libraryId);
    }
  };

  const handleRootFolderChange = async (_: any, data: any) => {
    const folderId = data.optionValue || "";
    const folderName = rootFolders.find((f) => f.id === folderId)?.name;
    setSelectedRootFolderId(folderId);
    setProjects([]);
    setSelectedProjectId("");
    setSubFolders([]);
    setSelectedSubFolderId("");
    setIsPrefilled(false);
    await loadProjects(selectedSiteId, selectedLibraryId, folderName);
  };

  const handleProjectChange = async (_: any, data: any) => {
    const projectId = data.optionValue;
    setSelectedProjectId(projectId);
    setSelectedSubFolderId("");
    setSubFolders([]);
    setIsPrefilled(false);
    await loadProjectSubFolders(selectedSiteId, projectId);
  };

  const handleSubFolderChange = (_: any, data: any) => {
    setSelectedSubFolderId(data.optionValue);
    setIsPrefilled(false);
  };

  const handleArchiveFolderChange = (_: any, data: any) => {
    setSelectedArchiveFolderId(data.optionValue || "");
  };

  // --- Opslaan ---
  const handleSave = useCallback(async () => {
    const item = getMailItem();
    if (!item || !selectedSiteId || !selectedLibraryId || !selectedProjectId || !selectedSubFolderId) {
      setErrorMessage("Selecteer een site, bibliotheek, project en map.");
      setSaveStatus("error");
      return;
    }

    setSaveStatus("saving");
    setErrorMessage("");

    try {
      const itemId = (item as any).itemId || item.itemId;
      const emlBlob = await getMailMimeContent(itemId);
      await uploadToSharePoint(selectedSiteId, selectedSubFolderId, fileName, emlBlob);

      const conversationId = getConversationId();
      const normalizedSubject = normalizeSubject(item.subject || "");
      const selectedSite = sites.find((s) => s.id === selectedSiteId);
      const selectedLibrary = libraries.find((l) => l.id === selectedLibraryId);
      const selectedProject = projects.find((p) => p.id === selectedProjectId);
      const selectedSubFolder = subFolders.find((f) => f.id === selectedSubFolderId);
      const selectedArchiveFolder = mailFolders.find((f) => f.id === selectedArchiveFolderId);

      if (conversationId) {
        await saveConversationHistory({
          conversationId,
          normalizedSubject,
          siteId: selectedSiteId,
          siteName: selectedSite?.displayName ?? "",
          libraryId: selectedLibraryId,
          libraryName: selectedLibrary?.name ?? "",
          projectFolderId: selectedProjectId,
          projectFolderName: selectedProject?.name ?? "",
          subFolderId: selectedSubFolderId,
          subFolderName: selectedSubFolder?.name ?? "",
          archiveMailFolderId: selectedArchiveFolderId || undefined,
          archiveMailFolderName: selectedArchiveFolder?.displayName,
          timestamp: Date.now(),
        });
      }

      if (selectedArchiveFolderId) {
        await moveMailToFolder(itemId, selectedArchiveFolderId);
      }

      setSaveStatus("success");
      setTimeout(() => setSaveStatus("idle"), 3000);
    } catch (error: any) {
      console.error("Opslaan mislukt:", error);
      setErrorMessage(error.message || "Er is een fout opgetreden.");
      setSaveStatus("error");
    }
  }, [
    selectedSiteId, selectedLibraryId, selectedProjectId, selectedSubFolderId,
    selectedArchiveFolderId, fileName, sites, libraries, projects, subFolders, mailFolders,
  ]);

  // --- Gefilterde projecten ---
  const filteredProjects = projectSearch
    ? projects.filter((p) => p.name.toLowerCase().includes(projectSearch.toLowerCase()))
    : projects;

  const canSave =
    selectedSiteId && selectedLibraryId && selectedProjectId && selectedSubFolderId && saveStatus !== "saving";

  // --- Render ---
  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text className={styles.title}>Mail naar SharePoint</Text>
        {isPrefilled && <Text className={styles.prefillBadge}>✓ Vooringevuld</Text>}
      </div>

      <div className={styles.form}>
        {/* Site */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>SharePoint Site</Text>
          <div className={styles.searchRow}>
            <Input
              placeholder="Zoek site..."
              value={siteSearch}
              onChange={(_, data) => setSiteSearch(data.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSiteSearch()}
              style={{ flexGrow: 1 }}
            />
            <Button icon={<Search20Regular />} onClick={handleSiteSearch} disabled={loadingSites} />
          </div>
          <Dropdown
            placeholder={loadingSites ? "Laden..." : "Selecteer site"}
            value={sites.find((s) => s.id === selectedSiteId)?.displayName ?? ""}
            selectedOptions={selectedSiteId ? [selectedSiteId] : []}
            onOptionSelect={handleSiteChange}
            disabled={loadingSites}
          >
            {sites.map((site) => (
              <Option key={site.id} value={site.id}>{site.displayName}</Option>
            ))}
          </Dropdown>
        </div>

        {/* Bibliotheek */}
        {selectedSiteId && (
          <div className={styles.fieldGroup}>
            <Text className={styles.label}>Bibliotheek</Text>
            <Dropdown
              placeholder={loadingLibraries ? "Laden..." : libraries.length === 0 ? "Geen bibliotheken gevonden" : "Selecteer bibliotheek"}
              value={libraries.find((l) => l.id === selectedLibraryId)?.name ?? ""}
              selectedOptions={selectedLibraryId ? [selectedLibraryId] : []}
              onOptionSelect={handleLibraryChange}
              disabled={loadingLibraries || libraries.length === 0}
            >
              {libraries.map((lib) => (
                <Option key={lib.id} value={lib.id}>{lib.name}</Option>
              ))}
            </Dropdown>
          </div>
        )}

        {/* Submap binnen bibliotheek (optioneel) */}
        {selectedLibraryId && rootFolders.length > 0 && (
          <div className={styles.fieldGroup}>
            <Text className={styles.label}>Submap in bibliotheek (optioneel)</Text>
            {loadingRootFolders ? (
              <Spinner size="tiny" label="Mappen laden..." />
            ) : (
              <Dropdown
                placeholder="Gebruik root van bibliotheek"
                value={rootFolders.find((f) => f.id === selectedRootFolderId)?.name ?? ""}
                selectedOptions={selectedRootFolderId ? [selectedRootFolderId] : []}
                onOptionSelect={handleRootFolderChange}
                clearable
              >
                {rootFolders.map((folder) => (
                  <Option key={folder.id} value={folder.id}>{folder.name}</Option>
                ))}
              </Dropdown>
            )}
          </div>
        )}

        {/* Project */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Project</Text>
          <Input
            placeholder="Filter op projectnummer (bijv. 24)..."
            value={projectSearch}
            onChange={(_, data) => setProjectSearch(data.value)}
            disabled={projects.length === 0 && !loadingProjects}
          />
          <Dropdown
            placeholder={
              loadingProjects ? "Laden..." :
              projects.length === 0 ? "Selecteer eerst een bibliotheek" :
              `${filteredProjects.length} project(en)`
            }
            value={projects.find((p) => p.id === selectedProjectId)?.name ?? ""}
            selectedOptions={selectedProjectId ? [selectedProjectId] : []}
            onOptionSelect={handleProjectChange}
            disabled={loadingProjects || projects.length === 0}
          >
            {filteredProjects.map((project) => (
              <Option key={project.id} value={project.id}>{project.name}</Option>
            ))}
          </Dropdown>
        </div>

        {/* Submap van project */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Map</Text>
          {loadingSubFolders ? (
            <Spinner size="tiny" label="Mappen laden..." />
          ) : (
            <Dropdown
              placeholder={
                !selectedProjectId ? "Selecteer eerst een project" :
                subFolders.length === 0 ? "Geen submappen gevonden" :
                "Selecteer map"
              }
              value={subFolders.find((f) => f.id === selectedSubFolderId)?.name ?? ""}
              selectedOptions={selectedSubFolderId ? [selectedSubFolderId] : []}
              onOptionSelect={handleSubFolderChange}
              disabled={!selectedProjectId || subFolders.length === 0}
            >
              {subFolders.map((folder) => (
                <Option key={folder.id} value={folder.id}>{folder.name}</Option>
              ))}
            </Dropdown>
          )}
        </div>

        <Divider />

        {/* Bestandsnaam */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Bestandsnaam</Text>
          <FileNameBuilder
            fields={fileNameFields}
            onChange={handleFileNameFieldsChange}
            preview={fileName || "..."}
          />
        </div>

        {/* Bijlagen */}
        <Checkbox
          checked={saveAttachments}
          onChange={(_, data) => setSaveAttachments(!!data.checked)}
          label="Bijlagen opslaan (in .eml)"
        />

        <Divider />

        {/* Archief map */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Archief map (optioneel)</Text>
          <Dropdown
            placeholder="Mail laten staan"
            value={mailFolders.find((f) => f.id === selectedArchiveFolderId)?.displayName ?? ""}
            selectedOptions={selectedArchiveFolderId ? [selectedArchiveFolderId] : []}
            onOptionSelect={handleArchiveFolderChange}
            clearable
          >
            {mailFolders.map((folder) => (
              <Option key={folder.id} value={folder.id}>{folder.displayName}</Option>
            ))}
          </Dropdown>
        </div>
      </div>

      {/* Footer */}
      <div className={styles.footer}>
        <Button
          className={styles.saveButton}
          appearance="primary"
          onClick={handleSave}
          disabled={!canSave}
        >
          {saveStatus === "saving" ? <Spinner size="tiny" label="Opslaan..." /> : "Opslaan in SharePoint"}
        </Button>

        {saveStatus === "success" && (
          <div className={`${styles.statusContainer} ${styles.statusSuccess}`}>
            <CheckmarkCircle24Filled />
            <Text>Opgeslagen!</Text>
          </div>
        )}
        {saveStatus === "error" && (
          <div className={`${styles.statusContainer} ${styles.statusError}`}>
            <DismissCircle24Filled />
            <Text>{errorMessage || "Er is een fout opgetreden."}</Text>
          </div>
        )}
      </div>
    </div>
  );
};
