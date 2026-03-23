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
  ProjectFolder,
  SubFolder,
  MailFolder,
  SaveStatus,
  ConversationHistory,
} from "../types";
import { APP_CONFIG, PROJECT_NUMBER_REGEX } from "../config";
import {
  searchSites,
  getSiteByHostname,
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
  fileName: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    wordBreak: "break-all" as const,
    ...shorthands.padding("4px", "8px"),
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.borderRadius("4px"),
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
// Component
// ============================================================

export const Taskpane: React.FC = () => {
  const styles = useStyles();

  // --- State ---
  const [sites, setSites] = useState<SharePointSite[]>([]);
  const [selectedSiteId, setSelectedSiteId] = useState<string>("");
  const [siteSearch, setSiteSearch] = useState<string>("");

  const [projects, setProjects] = useState<ProjectFolder[]>([]);
  const [selectedProjectId, setSelectedProjectId] = useState<string>("");
  const [projectSearch, setProjectSearch] = useState<string>("");

  const [subFolders, setSubFolders] = useState<SubFolder[]>([]);
  const [selectedSubFolderId, setSelectedSubFolderId] = useState<string>("");

  const [mailFolders, setMailFolders] = useState<MailFolder[]>([]);
  const [selectedArchiveFolderId, setSelectedArchiveFolderId] = useState<string>("");

  const [saveAttachments, setSaveAttachments] = useState<boolean>(true);
  const [fileName, setFileName] = useState<string>("");
  const [saveStatus, setSaveStatus] = useState<SaveStatus>("idle");
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [isPrefilled, setIsPrefilled] = useState<boolean>(false);

  const [loadingSites, setLoadingSites] = useState(false);
  const [loadingProjects, setLoadingProjects] = useState(false);
  const [loadingSubFolders, setLoadingSubFolders] = useState(false);

  // --- Helpers ---
  const getMailItem = () => Office.context.mailbox.item;

  const getConversationId = (): string | undefined => {
    return getMailItem()?.conversationId;
  };

  // --- Initialisatie ---
  useEffect(() => {
    initializeTaskpane();
  }, []);

  const initializeTaskpane = async () => {
    try {
      // Genereer bestandsnaam
      const item = getMailItem();
      if (item) {
        const subject = item.subject || "Geen onderwerp";
        const dateReceived = item.dateTimeCreated
          ? new Date(item.dateTimeCreated.toString())
          : new Date();
        setFileName(generateEmailFileName(subject, dateReceived));
      }

      // Laad standaard site
      await loadDefaultSite();

      // Laad mail folders voor archief selector
      loadMailFolders();

      // Check roaming settings voor smart prefill
      const conversationId = getConversationId();
      if (conversationId) {
        const history = getHistoryForConversation(conversationId);
        if (history) {
          applyHistory(history);
        }
      }
    } catch (error) {
      console.error("Initialisatie fout:", error);
    }
  };

  const loadDefaultSite = async () => {
    setLoadingSites(true);
    try {
      const defaultSite = await getSiteByHostname(
        APP_CONFIG.defaultSiteHostname
      );
      setSites([defaultSite]);
      setSelectedSiteId(defaultSite.id);
      // Laad projecten voor de standaard site
      await loadProjects(defaultSite.id);
    } catch (error) {
      console.error("Kan standaard site niet laden:", error);
      // Probeer te zoeken
      const results = await searchSites(APP_CONFIG.defaultSiteHostname);
      setSites(results);
    } finally {
      setLoadingSites(false);
    }
  };

  const loadProjects = async (siteId: string) => {
    setLoadingProjects(true);
    try {
      const basePath = getBasePathForSite(siteId);
      const folders = await getProjectFolders(siteId, basePath);
      setProjects(folders);
    } catch (error) {
      console.error("Kan projecten niet laden:", error);
      setProjects([]);
    } finally {
      setLoadingProjects(false);
    }
  };

  const loadSubFolders = async (siteId: string, projectId: string) => {
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
      const folders = await getMailFolders();
      setMailFolders(folders);
    } catch (error) {
      console.error("Kan mail folders niet laden:", error);
    }
  };

  const getBasePathForSite = (siteId: string): string => {
    const config = APP_CONFIG.siteConfigs.find((c) => c.siteId === siteId);
    return config?.basePath ?? APP_CONFIG.defaultBasePath;
  };

  // --- Smart Prefill ---
  const applyHistory = async (history: ConversationHistory) => {
    setIsPrefilled(true);

    // Stel site in
    setSelectedSiteId(history.siteId);

    // Laad projecten en stel project in
    await loadProjects(history.siteId);
    setSelectedProjectId(history.projectFolderId);

    // Laad submappen en stel submap in
    await loadSubFolders(history.siteId, history.projectFolderId);
    setSelectedSubFolderId(history.subFolderId);

    // Archief map
    if (history.archiveMailFolderId) {
      setSelectedArchiveFolderId(history.archiveMailFolderId);
    }
  };

  // --- Site zoeken ---
  const handleSiteSearch = async () => {
    if (!siteSearch.trim()) return;
    setLoadingSites(true);
    try {
      const results = await searchSites(siteSearch);
      setSites(results);
    } catch (error) {
      console.error("Site zoeken mislukt:", error);
    } finally {
      setLoadingSites(false);
    }
  };

  // --- Event handlers ---
  const handleSiteChange = async (_: any, data: any) => {
    const siteId = data.optionValue;
    setSelectedSiteId(siteId);
    setSelectedProjectId("");
    setSelectedSubFolderId("");
    setSubFolders([]);
    setIsPrefilled(false);
    await loadProjects(siteId);
  };

  const handleProjectChange = async (_: any, data: any) => {
    const projectId = data.optionValue;
    setSelectedProjectId(projectId);
    setSelectedSubFolderId("");
    setIsPrefilled(false);
    await loadSubFolders(selectedSiteId, projectId);
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
    if (!item || !selectedSiteId || !selectedProjectId || !selectedSubFolderId) {
      setErrorMessage("Selecteer een site, project en map.");
      setSaveStatus("error");
      return;
    }

    setSaveStatus("saving");
    setErrorMessage("");

    try {
      // 1. Haal de e-mail op als .eml (MIME content)
      const itemId = (item as any).itemId || item.itemId;
      const emlBlob = await getMailMimeContent(itemId);

      // 2. Upload naar SharePoint
      await uploadToSharePoint(
        selectedSiteId,
        selectedSubFolderId,
        fileName,
        emlBlob
      );

      // 3. Sla conversatie history op (smart prefill)
      const conversationId = getConversationId();
      if (conversationId) {
        const selectedSite = sites.find((s) => s.id === selectedSiteId);
        const selectedProject = projects.find((p) => p.id === selectedProjectId);
        const selectedSubFolder = subFolders.find((f) => f.id === selectedSubFolderId);
        const selectedArchiveFolder = mailFolders.find(
          (f) => f.id === selectedArchiveFolderId
        );

        await saveConversationHistory({
          conversationId,
          siteId: selectedSiteId,
          siteName: selectedSite?.displayName ?? "",
          projectFolderId: selectedProjectId,
          projectFolderName: selectedProject?.name ?? "",
          subFolderId: selectedSubFolderId,
          subFolderName: selectedSubFolder?.name ?? "",
          archiveMailFolderId: selectedArchiveFolderId || undefined,
          archiveMailFolderName: selectedArchiveFolder?.displayName,
          timestamp: Date.now(),
        });
      }

      // 4. Verplaats mail naar archief (als geselecteerd)
      if (selectedArchiveFolderId) {
        const itemId = (item as any).itemId || item.itemId;
        await moveMailToFolder(itemId, selectedArchiveFolderId);
      }

      setSaveStatus("success");

      // Reset status na 3 seconden
      setTimeout(() => setSaveStatus("idle"), 3000);
    } catch (error: any) {
      console.error("Opslaan mislukt:", error);
      setErrorMessage(error.message || "Er is een fout opgetreden.");
      setSaveStatus("error");
    }
  }, [
    selectedSiteId,
    selectedProjectId,
    selectedSubFolderId,
    selectedArchiveFolderId,
    fileName,
    sites,
    projects,
    subFolders,
    mailFolders,
  ]);

  // --- Gefilterde projecten (zoekfunctie) ---
  const filteredProjects = projectSearch
    ? projects.filter((p) =>
        p.name.toLowerCase().includes(projectSearch.toLowerCase())
      )
    : projects;

  // --- Render ---
  const canSave =
    selectedSiteId && selectedProjectId && selectedSubFolderId && saveStatus !== "saving";

  return (
    <div className={styles.root}>
      {/* Header */}
      <div className={styles.header}>
        <Text className={styles.title}>Mail naar SharePoint</Text>
        {isPrefilled && (
          <Text className={styles.prefillBadge}>✓ Vooringevuld</Text>
        )}
      </div>

      {/* Form */}
      <div className={styles.form}>
        {/* Site Dropdown */}
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
            <Button
              icon={<Search20Regular />}
              onClick={handleSiteSearch}
              disabled={loadingSites}
            />
          </div>
          <Dropdown
            placeholder={loadingSites ? "Laden..." : "Selecteer site"}
            value={
              sites.find((s) => s.id === selectedSiteId)?.displayName ?? ""
            }
            selectedOptions={selectedSiteId ? [selectedSiteId] : []}
            onOptionSelect={handleSiteChange}
            disabled={loadingSites}
          >
            {sites.map((site) => (
              <Option key={site.id} value={site.id}>
                {site.displayName}
              </Option>
            ))}
          </Dropdown>
        </div>

        {/* Project Dropdown */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Project</Text>
          <Input
            placeholder="Zoek project (bijv. 24-001)..."
            value={projectSearch}
            onChange={(_, data) => setProjectSearch(data.value)}
          />
          <Dropdown
            placeholder={
              loadingProjects
                ? "Laden..."
                : projects.length === 0
                ? "Selecteer eerst een site"
                : "Selecteer project"
            }
            value={
              projects.find((p) => p.id === selectedProjectId)?.name ?? ""
            }
            selectedOptions={selectedProjectId ? [selectedProjectId] : []}
            onOptionSelect={handleProjectChange}
            disabled={loadingProjects || projects.length === 0}
          >
            {filteredProjects.map((project) => (
              <Option key={project.id} value={project.id}>
                {project.name}
              </Option>
            ))}
          </Dropdown>
        </div>

        {/* Submap Dropdown */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Map</Text>
          {loadingSubFolders ? (
            <Spinner size="tiny" label="Submappen laden..." />
          ) : (
            <Dropdown
              placeholder={
                subFolders.length === 0
                  ? "Selecteer eerst een project"
                  : "Selecteer map"
              }
              value={
                subFolders.find((f) => f.id === selectedSubFolderId)?.name ?? ""
              }
              selectedOptions={
                selectedSubFolderId ? [selectedSubFolderId] : []
              }
              onOptionSelect={handleSubFolderChange}
              disabled={subFolders.length === 0}
            >
              {subFolders.map((folder) => (
                <Option key={folder.id} value={folder.id}>
                  {folder.name}
                </Option>
              ))}
            </Dropdown>
          )}
        </div>

        <Divider />

        {/* Bestandsnaam */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Bestandsnaam</Text>
          <Text className={styles.fileName}>{fileName || "..."}</Text>
        </div>

        {/* Bijlagen checkbox */}
        <Checkbox
          checked={saveAttachments}
          onChange={(_, data) => setSaveAttachments(!!data.checked)}
          label="Bijlagen opslaan (in .eml)"
        />

        <Divider />

        {/* Archief Map Selector */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Archief map (optioneel)</Text>
          <Dropdown
            placeholder="Mail laten staan"
            value={
              mailFolders.find((f) => f.id === selectedArchiveFolderId)
                ?.displayName ?? ""
            }
            selectedOptions={
              selectedArchiveFolderId ? [selectedArchiveFolderId] : []
            }
            onOptionSelect={handleArchiveFolderChange}
            clearable
          >
            {mailFolders.map((folder) => (
              <Option key={folder.id} value={folder.id}>
                {folder.displayName}
              </Option>
            ))}
          </Dropdown>
        </div>
      </div>

      {/* Footer - Save Button & Status */}
      <div className={styles.footer}>
        <Button
          className={styles.saveButton}
          appearance="primary"
          onClick={handleSave}
          disabled={!canSave}
        >
          {saveStatus === "saving" ? (
            <Spinner size="tiny" label="Opslaan..." />
          ) : (
            "Opslaan in SharePoint"
          )}
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
