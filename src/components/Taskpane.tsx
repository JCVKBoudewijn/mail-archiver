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
  ToggleButton,
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
  OrgConfig,
} from "../types";
import { APP_CONFIG, AUTO_SUBFOLDER_NAME } from "../config";
import { FileNameBuilder } from "./FileNameBuilder";
import {
  getUserEmail,
  getSiteByHostname,
  getLibraryByName,
  searchSites,
  getLibraries,
  getRootFolders,
  getProjectFolders,
  getSubFolders,
  getSubFoldersInDrive,
  getOrCreateSubFolder,
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
  orgBadge: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorPaletteBlueForeground2,
    fontStyle: "italic" as const,
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
  toggleRow: {
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

type WorkType = "werken" | "projecten";

function normalizeSubject(subject: string): string {
  return subject
    .replace(/^(re|fw|fwd|tr|antw):\s*/gi, "")
    .trim()
    .toLowerCase();
}

function detectOrg(email: string): OrgConfig | null {
  const domain = email.split("@")[1]?.toLowerCase();
  if (!domain) return null;
  return APP_CONFIG.orgConfigs.find((org) =>
    org.emailDomains.some((d) => d.toLowerCase() === domain)
  ) || null;
}

// ============================================================
// Component
// ============================================================

export const Taskpane: React.FC = () => {
  const styles = useStyles();

  // Org detectie
  const [detectedOrg, setDetectedOrg] = useState<OrgConfig | null>(null);
  const [orgLoading, setOrgLoading] = useState(true);
  const [workType, setWorkType] = useState<WorkType>("werken");

  // Org-mode state
  const [orgSiteId, setOrgSiteId] = useState<string>("");
  const [orgDriveId, setOrgDriveId] = useState<string>("");

  // Fallback: Site
  const [sites, setSites] = useState<SharePointSite[]>([]);
  const [selectedSiteId, setSelectedSiteId] = useState<string>("");
  const [siteSearch, setSiteSearch] = useState<string>("");
  const [loadingSites, setLoadingSites] = useState(false);

  // Fallback: Bibliotheek
  const [libraries, setLibraries] = useState<Library[]>([]);
  const [selectedLibraryId, setSelectedLibraryId] = useState<string>("");
  const [loadingLibraries, setLoadingLibraries] = useState(false);

  // Fallback: Submap binnen bibliotheek
  const [rootFolders, setRootFolders] = useState<SubFolder[]>([]);
  const [selectedRootFolderId, setSelectedRootFolderId] = useState<string>("");
  const [loadingRootFolders, setLoadingRootFolders] = useState(false);

  // Project (gedeeld)
  const [projects, setProjects] = useState<ProjectFolder[]>([]);
  const [selectedProjectId, setSelectedProjectId] = useState<string>("");
  const [projectSearch, setProjectSearch] = useState<string>("");
  const [loadingProjects, setLoadingProjects] = useState(false);

  // Submap van project (gedeeld)
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

  // Gedeelde mailbox
  const [sharedMailboxUser, setSharedMailboxUser] = useState<string | undefined>(undefined);

  // --- Helpers ---
  const getMailItem = () => Office.context.mailbox.item;
  const getConversationId = (): string | undefined => getMailItem()?.conversationId;

  // Effectieve IDs (org-mode of fallback)
  const activeSiteId = detectedOrg ? orgSiteId : selectedSiteId;
  const activeDriveId = detectedOrg ? orgDriveId : selectedLibraryId;

  // --- Initialisatie ---
  useEffect(() => {
    initializeTaskpane();
  }, []);

  const initializeTaskpane = async () => {
    try {
      let detectedSharedUser: string | undefined;

      // Mail metadata laden
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

        // Detecteer gedeelde mailbox (Office.js 1.8+, stille fallback op oudere versies)
        if (typeof (item as any).getSharedPropertiesAsync === "function") {
          await new Promise<void>((resolve) => {
            (item as any).getSharedPropertiesAsync((result: any) => {
              if (
                result.status === Office.AsyncResultStatus.Succeeded &&
                result.value?.targetMailbox
              ) {
                detectedSharedUser = result.value.targetMailbox;
                setSharedMailboxUser(detectedSharedUser);
              }
              resolve();
            });
          });
        }
      }

      // Org detectie
      const email = await getUserEmail();
      const org = detectOrg(email);
      setDetectedOrg(org);

      if (org) {
        await initOrgMode(org);
      } else {
        await initFallbackMode();
      }

      loadMailFolders(detectedSharedUser);
    } catch (error) {
      console.error("Initialisatie fout:", error);
      setOrgLoading(false);
    }
  };

  // --- Org-mode initialisatie ---
  const initOrgMode = async (org: OrgConfig) => {
    setOrgLoading(true);
    try {
      // Site ophalen
      const site = await getSiteByHostname(APP_CONFIG.tenantHostname, org.siteUrl);
      setOrgSiteId(site.id);

      // Bibliotheek ophalen op basis van standaard workType (werken)
      const libConfig = org.werken;
      const lib = await getLibraryByName(site.id, libConfig.libraryName);
      if (lib) {
        setOrgDriveId(lib.id);
        await loadProjects(site.id, lib.id, libConfig.subPath);
      }
    } catch (error) {
      console.error("Org-mode initialisatie mislukt:", error);
    } finally {
      setOrgLoading(false);
    }
  };

  // --- Work type toggle ---
  const handleWorkTypeChange = async (newType: WorkType) => {
    if (!detectedOrg || newType === workType) return;
    setWorkType(newType);
    setProjects([]);
    setSelectedProjectId("");
    setSubFolders([]);
    setSelectedSubFolderId("");

    const libConfig = newType === "werken" ? detectedOrg.werken : detectedOrg.projecten;
    if (!libConfig) return; // projecten niet beschikbaar

    try {
      setLoadingProjects(true);
      const lib = await getLibraryByName(orgSiteId, libConfig.libraryName);
      if (lib) {
        setOrgDriveId(lib.id);
        await loadProjects(orgSiteId, lib.id, libConfig.subPath);
      }
    } catch (error) {
      console.error("Bibliotheek wisselen mislukt:", error);
    } finally {
      setLoadingProjects(false);
    }
  };

  // --- Fallback-mode initialisatie ---
  const initFallbackMode = async () => {
    setOrgLoading(false);
    setLoadingSites(true);
    try {
      const results = await searchSites("*");
      setSites(results);

      // Prefill via conversatie-history
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
      }
    } catch (error) {
      console.error("Kan sites niet laden:", error);
    } finally {
      setLoadingSites(false);
    }
  };

  // --- Gedeelde functies ---
  const loadProjects = async (siteId: string, libraryId: string, subPath?: string) => {
    setLoadingProjects(true);
    try {
      const folders = await getProjectFolders(siteId, libraryId, subPath);
      setProjects(folders);
    } catch (error: any) {
      console.error("Kan projecten niet laden:", error?.message || error);
      setProjects([]);
    } finally {
      setLoadingProjects(false);
    }
  };

  const handleProjectChange = async (_: any, data: any) => {
    const projectId = data.optionValue;
    setSelectedProjectId(projectId);
    setSelectedSubFolderId("");
    setSubFolders([]);
    setIsPrefilled(false);
    setLoadingSubFolders(true);
    try {
      const folders = detectedOrg
        ? await getSubFoldersInDrive(orgSiteId, orgDriveId, projectId)
        : await getSubFolders(activeSiteId, projectId);
      setSubFolders(folders);
    } catch (error) {
      console.error("Kan submappen niet laden:", error);
      setSubFolders([]);
    } finally {
      setLoadingSubFolders(false);
    }
  };

  const handleSubFolderChange = (_: any, data: any) => {
    setSelectedSubFolderId(data.optionValue);
    setIsPrefilled(false);
  };

  const loadMailFolders = async (mailboxUser?: string) => {
    try {
      setMailFolders(await getMailFolders(mailboxUser));
    } catch (error) {
      console.error("Kan mail folders niet laden:", error);
    }
  };

  // --- Fallback event handlers ---
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

    const rootFolderList = await loadRootFoldersForLibrary(selectedSiteId, libraryId);

    if (rootFolderList.length === 1) {
      setSelectedRootFolderId(rootFolderList[0].id);
      await loadProjects(selectedSiteId, libraryId, rootFolderList[0].name);
    } else {
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

  const handleArchiveFolderChange = (_: any, data: any) => {
    setSelectedArchiveFolderId(data.optionValue || "");
  };

  // --- Smart Prefill (fallback) ---
  const applyHistory = async (history: ConversationHistory, availableSites: SharePointSite[]) => {
    if (!availableSites.find((s) => s.id === history.siteId)) return;
    setIsPrefilled(true);
    setSelectedSiteId(history.siteId);

    const libs = await loadLibrariesForSite(history.siteId);
    if (!libs.find((l) => l.id === history.libraryId)) return;
    setSelectedLibraryId(history.libraryId);

    const rootFolderList = await loadRootFoldersForLibrary(history.siteId, history.libraryId);

    const matchedRoot = rootFolderList.find((f) =>
      history.projectFolderName?.startsWith(f.name) === false &&
      history.libraryName === f.name
    );
    const subFolderId = matchedRoot?.id || "";
    setSelectedRootFolderId(subFolderId);

    await loadProjects(history.siteId, history.libraryId, subFolderId || undefined);
    setSelectedProjectId(history.projectFolderId);

    setLoadingSubFolders(true);
    try {
      const folders = await getSubFolders(history.siteId, history.projectFolderId);
      setSubFolders(folders);
    } catch { setSubFolders([]); } finally { setLoadingSubFolders(false); }
    setSelectedSubFolderId(history.subFolderId);

    if (history.archiveMailFolderId) setSelectedArchiveFolderId(history.archiveMailFolderId);
  };

  // --- Bestandsnaam ---
  const handleFileNameFieldsChange = async (fields: FileNameField[]) => {
    setFileNameFields(fields);
    setFileName(generateEmailFileName(mailSubject, mailDate, fields, mailSender, mailRecipient));
    await saveFileNameConfig({ fields }).catch(console.error);
  };

  // --- Opslaan ---
  const handleSave = useCallback(async () => {
    const item = getMailItem();
    if (!item || !activeSiteId || !activeDriveId || !selectedProjectId) {
      setErrorMessage("Selecteer een project.");
      setSaveStatus("error");
      return;
    }

    setSaveStatus("saving");
    setErrorMessage("");

    try {
      const itemId = (item as any).itemId || item.itemId;
      const emlBlob = await getMailMimeContent(itemId, sharedMailboxUser);

      let targetFolderId: string;

      if (detectedOrg) {
        // Org-mode: auto-create Correspondentie en upload daarin
        const corrFolder = await getOrCreateSubFolder(
          orgSiteId, orgDriveId, selectedProjectId, AUTO_SUBFOLDER_NAME
        );
        targetFolderId = corrFolder.id;

        // Update subfolders weergave als Correspondentie net aangemaakt is
        if (!subFolders.find((f) => f.id === corrFolder.id)) {
          setSubFolders((prev) => [...prev, corrFolder].sort((a, b) => a.name.localeCompare(b.name)));
        }
        setSelectedSubFolderId(corrFolder.id);
      } else {
        // Fallback: gebruik geselecteerde submap of projectmap
        targetFolderId = selectedSubFolderId || selectedProjectId;
      }

      await uploadToSharePoint(activeSiteId, targetFolderId, fileName, emlBlob);

      // Conversatie-history opslaan
      const conversationId = getConversationId();
      const normalizedSubject = normalizeSubject(item.subject || "");

      if (conversationId) {
        const selectedProject = projects.find((p) => p.id === selectedProjectId);

        await saveConversationHistory({
          conversationId,
          normalizedSubject,
          siteId: activeSiteId,
          siteName: detectedOrg?.name ?? sites.find((s) => s.id === activeSiteId)?.displayName ?? "",
          libraryId: activeDriveId,
          libraryName: detectedOrg
            ? (workType === "werken" ? detectedOrg.werken.libraryName : detectedOrg.projecten?.libraryName ?? "")
            : libraries.find((l) => l.id === activeDriveId)?.name ?? "",
          projectFolderId: selectedProjectId,
          projectFolderName: selectedProject?.name ?? "",
          subFolderId: selectedSubFolderId,
          subFolderName: subFolders.find((f) => f.id === selectedSubFolderId)?.name ?? "",
          archiveMailFolderId: selectedArchiveFolderId || undefined,
          archiveMailFolderName: mailFolders.find((f) => f.id === selectedArchiveFolderId)?.displayName,
          timestamp: Date.now(),
        });
      }

      if (selectedArchiveFolderId) {
        await moveMailToFolder(itemId, selectedArchiveFolderId, sharedMailboxUser);
      }

      setSaveStatus("success");
      setTimeout(() => setSaveStatus("idle"), 3000);
    } catch (error: any) {
      console.error("Opslaan mislukt:", error);
      setErrorMessage(error.message || "Er is een fout opgetreden.");
      setSaveStatus("error");
    }
  }, [
    activeSiteId, activeDriveId, selectedProjectId, selectedSubFolderId,
    selectedArchiveFolderId, fileName, detectedOrg, orgSiteId, orgDriveId,
    workType, sites, libraries, projects, subFolders, mailFolders, sharedMailboxUser,
  ]);

  // --- Gefilterde projecten ---
  const filteredProjects = projectSearch
    ? projects.filter((p) => p.name.toLowerCase().includes(projectSearch.toLowerCase()))
    : projects;

  const canSave =
    activeSiteId && activeDriveId && selectedProjectId && saveStatus !== "saving";

  // --- Loading state ---
  if (orgLoading) {
    return (
      <div className={styles.root}>
        <Spinner label="Organisatie detecteren..." />
      </div>
    );
  }

  // --- Render ---
  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text className={styles.title}>Mail naar SharePoint</Text>
        {detectedOrg && <Text className={styles.orgBadge}>{detectedOrg.name}</Text>}
        {!detectedOrg && isPrefilled && <Text className={styles.prefillBadge}>Vooringevuld</Text>}
      </div>

      <div className={styles.form}>
        {/* ============ ORG-MODE ============ */}
        {detectedOrg && (
          <>
            {/* Werken / Projecten toggle */}
            <div className={styles.fieldGroup}>
              <Text className={styles.label}>Type</Text>
              <div className={styles.toggleRow}>
                <ToggleButton
                  checked={workType === "werken"}
                  onClick={() => handleWorkTypeChange("werken")}
                  appearance={workType === "werken" ? "primary" : "outline"}
                  style={{ flexGrow: 1 }}
                >
                  Werken
                </ToggleButton>
                <ToggleButton
                  checked={workType === "projecten"}
                  onClick={() => handleWorkTypeChange("projecten")}
                  appearance={workType === "projecten" ? "primary" : "outline"}
                  disabled={!detectedOrg.projecten}
                  title={!detectedOrg.projecten ? "Nog niet beschikbaar" : undefined}
                  style={{ flexGrow: 1 }}
                >
                  Projecten
                </ToggleButton>
              </div>
            </div>
          </>
        )}

        {/* ============ FALLBACK-MODE ============ */}
        {!detectedOrg && (
          <>
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

            {/* Submap binnen bibliotheek */}
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
          </>
        )}

        {/* ============ GEDEELD: Project selectie ============ */}
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
              projects.length === 0 ? (detectedOrg ? "Laden..." : "Selecteer eerst een bibliotheek") :
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
        {!detectedOrg && (
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
        )}

        {detectedOrg && selectedProjectId && (
          <div className={styles.fieldGroup}>
            <Text className={styles.label}>Opslaan in</Text>
            <Text style={{ fontSize: tokens.fontSizeBase300, color: tokens.colorNeutralForeground2 }}>
              {projects.find((p) => p.id === selectedProjectId)?.name} / {AUTO_SUBFOLDER_NAME}
            </Text>
          </div>
        )}

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
