import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Button,
  Combobox,
  Dropdown,
  Option,
  Checkbox,
  Spinner,
  Text,
  Divider,
  makeStyles,
  tokens,
  ToggleButton,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  DismissCircle24Filled,
} from "@fluentui/react-icons";

import type {
  ProjectFolder,
  Library,
  MailFolder,
  SaveStatus,
  FileNameField,
  OrgConfig,
  WorkType,
} from "../types";
import { APP_CONFIG, AUTO_SUBFOLDER_NAME, PROJECT_NUMBER_REGEX } from "../config";
import { FileNameBuilder } from "./FileNameBuilder";
import {
  getSiteByHostname,
  getLibraries,
  getLibraryByName,
  getProjectFolders,
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
    padding: "16px",
    boxSizing: "border-box",
    fontFamily: tokens.fontFamilyBase,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    display: "flex",
    alignItems: "center",
    marginBottom: "16px",
    gap: "8px",
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
    gap: "12px",
    flexGrow: 1,
    overflowY: "auto" as const,
  },
  fieldGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    textTransform: "uppercase" as const,
    letterSpacing: "0.5px",
  },
  toggleRow: {
    display: "flex",
    gap: "4px",
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
    gap: "8px",
    padding: "12px",
    borderRadius: "8px",
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

function detectOrg(email: string): OrgConfig | null {
  const domain = email.split("@")[1]?.toLowerCase();
  if (!domain) return null;
  return APP_CONFIG.orgConfigs.find((org) =>
    org.emailDomains.some((d) => d.toLowerCase() === domain)
  ) || null;
}

function extractProjectNumber(subject: string): string {
  const match = subject.match(PROJECT_NUMBER_REGEX);
  return match ? match[0] : "";
}

// ============================================================
// Component
// ============================================================

export const Taskpane: React.FC = () => {
  const styles = useStyles();

  // Org
  const [detectedOrg, setDetectedOrg] = useState<OrgConfig | null>(null);
  const [orgError, setOrgError] = useState<string>("");
  const [orgLoading, setOrgLoading] = useState(true);
  const [workType, setWorkType] = useState<WorkType>("werken");

  // Actieve SharePoint-locatie (ingesteld vanuit org-config)
  const [siteId, setSiteId] = useState<string>("");
  const [driveId, setDriveId] = useState<string>("");
  const siteIdRef = useRef<string>("");

  // Fallback: bibliotheek handmatig kiezen als auto-detectie mislukt
  const [libraries, setLibraries] = useState<Library[]>([]);
  const [selectedLibraryId, setSelectedLibraryId] = useState<string>("");

  // Projecten
  const [projects, setProjects] = useState<ProjectFolder[]>([]);
  const [selectedProjectId, setSelectedProjectId] = useState<string>("");
  const [projectInput, setProjectInput] = useState<string>("");
  const [loadingProjects, setLoadingProjects] = useState(false);

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

  // Gedeelde mailbox
  const [sharedMailboxUser, setSharedMailboxUser] = useState<string | undefined>(undefined);

  const getMailItem = () => Office.context.mailbox.item;

  // Gefilterde projecten op basis van wat de gebruiker typt
  const filteredProjects = projectInput
    ? projects.filter((p) => p.name.toLowerCase().includes(projectInput.toLowerCase()))
    : projects;

  // --- Initialisatie ---
  useEffect(() => {
    initializeTaskpane();
  }, []);

  const initializeTaskpane = async () => {
    try {
      const item = getMailItem();
      if (!item) return;

      // Mail metadata
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

      // Projectnummer uit onderwerp pre-invullen
      const detectedNumber = extractProjectNumber(subject);
      if (detectedNumber) setProjectInput(detectedNumber);

      // Gedeelde mailbox detecteren
      let mailboxEmail: string | undefined;
      if (typeof (item as any).getSharedPropertiesAsync === "function") {
        mailboxEmail = await new Promise<string | undefined>((resolve) => {
          (item as any).getSharedPropertiesAsync((result: any) => {
            if (
              result.status === Office.AsyncResultStatus.Succeeded &&
              result.value?.targetMailbox
            ) {
              resolve(result.value.targetMailbox);
            } else {
              resolve(undefined);
            }
          });
        });
        if (mailboxEmail) setSharedMailboxUser(mailboxEmail);
      }

      // Org detecteren op basis van actieve mailbox (gedeeld of eigen)
      // Office.js heeft het e-mailadres lokaal beschikbaar — geen Graph call nodig
      const ownEmail = Office.context.mailbox.userProfile.emailAddress;
      const emailForDetection = mailboxEmail || ownEmail;
      const org = detectOrg(emailForDetection);
      setDetectedOrg(org);

      if (!org) {
        const domain = emailForDetection.split("@")[1] || "onbekend";
        setOrgError(`Geen configuratie gevonden voor @${domain}. Neem contact op met de beheerder.`);
        setOrgLoading(false);
        return;
      }

      // Org-mode initialiseren
      await initOrgMode(org, item.conversationId, normalizeSubject(subject));

      // Mail-archiefmappen laden
      loadMailFolders(mailboxEmail);
    } catch (error) {
      console.error("Initialisatie fout:", error);
      setOrgError("Er is een fout opgetreden bij het laden.");
      setOrgLoading(false);
    }
  };

  const initOrgMode = async (
    org: OrgConfig,
    conversationId?: string,
    normalizedSubject?: string
  ) => {
    try {
      const site = await getSiteByHostname(APP_CONFIG.tenantHostname, org.siteUrl);
      setSiteId(site.id);
      siteIdRef.current = site.id;

      const libConfig = org.werken;
      const lib = await getLibraryByName(site.id, libConfig.libraryName);

      if (!lib) {
        // Bibliotheek niet gevonden op naam — toon fallback dropdown met alle bibliotheken
        console.warn(`Bibliotheek "${libConfig.libraryName}" niet gevonden, fallback naar handmatige selectie`);
        const allLibs = await getLibraries(site.id);
        setLibraries(allLibs);
        return; // driveId blijft leeg → UI toont bibliotheek-dropdown
      }

      await loadLibraryAndProjects(site.id, lib.id, libConfig.subPath, conversationId, normalizedSubject);
    } catch (error) {
      console.error("Org-mode initialisatie mislukt:", error);
      setOrgError("Kan SharePoint site niet bereiken.");
    } finally {
      setOrgLoading(false);
    }
  };

  const loadLibraryAndProjects = async (
    resolvedSiteId: string,
    resolvedDriveId: string,
    subPath?: string,
    conversationId?: string,
    normalizedSubject?: string
  ) => {
    setDriveId(resolvedDriveId);
    const folders = await getProjectFolders(resolvedSiteId, resolvedDriveId, subPath);
    setProjects(folders);

    const history =
      (conversationId ? getHistoryForConversation(conversationId) : undefined) ??
      (normalizedSubject
        ? loadConversationHistory().find((h) => h.normalizedSubject === normalizedSubject)
        : undefined);

    if (history && folders.find((p) => p.id === history.projectFolderId)) {
      setSelectedProjectId(history.projectFolderId);
      setProjectInput(history.projectFolderName);
      if (history.archiveMailFolderId) setSelectedArchiveFolderId(history.archiveMailFolderId);
      if (history.fileNameFields) {
        setFileNameFields(history.fileNameFields);
        setFileName(generateEmailFileName(mailSubject, mailDate, history.fileNameFields, mailSender, mailRecipient));
      }
    }
  };

  const handleFallbackLibraryChange = async (_: any, data: any) => {
    const libId = data.optionValue;
    setSelectedLibraryId(libId);
    setProjects([]);
    setSelectedProjectId("");
    setProjectInput("");
    setLoadingProjects(true);
    try {
      await loadLibraryAndProjects(siteId, libId);
    } catch (error) {
      console.error("Projecten laden mislukt:", error);
    } finally {
      setLoadingProjects(false);
    }
  };

  // --- Werken / Projecten toggle ---
  const handleWorkTypeChange = async (newType: WorkType) => {
    if (!detectedOrg || newType === workType) return;

    const libConfig = newType === "werken" ? detectedOrg.werken : detectedOrg.projecten;
    if (!libConfig) return;

    setWorkType(newType);
    setProjects([]);
    setSelectedProjectId("");
    setProjectInput("");
    setLoadingProjects(true);

    // Gebruik ref zodat we altijd de actuele siteId hebben (niet stale state)
    const currentSiteId = siteIdRef.current;

    try {
      const lib = await getLibraryByName(currentSiteId, libConfig.libraryName);
      if (lib) {
        setDriveId(lib.id);
        const folders = await getProjectFolders(currentSiteId, lib.id, libConfig.subPath);
        setProjects(folders);
      } else {
        console.warn(`Bibliotheek "${libConfig.libraryName}" niet gevonden`);
      }
    } catch (error) {
      console.error("Bibliotheek wisselen mislukt:", error);
    } finally {
      setLoadingProjects(false);
    }
  };

  const loadMailFolders = async (mailboxUser?: string) => {
    try {
      setMailFolders(await getMailFolders(mailboxUser));
    } catch (error) {
      console.error("Kan mail folders niet laden:", error);
    }
  };

  // --- Project combobox handlers ---
  const handleProjectSelect = async (_: any, data: any) => {
    setSelectedProjectId(data.optionValue ?? "");
    setProjectInput(data.optionText ?? "");
  };

  const handleProjectInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setProjectInput(e.target.value);
    setSelectedProjectId("");
  };

  const handleArchiveFolderChange = (_: any, data: any) => {
    setSelectedArchiveFolderId(data.optionValue || "");
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
    if (!item || !siteId || !driveId || !selectedProjectId) {
      setErrorMessage("Selecteer een project.");
      setSaveStatus("error");
      return;
    }

    setSaveStatus("saving");
    setErrorMessage("");

    try {
      // Office.js geeft EWS-formaat ID terug; Graph API verwacht REST-formaat.
      // convertToRestId is alleen beschikbaar in klassiek Outlook desktop — in
      // web geeft item.itemId al een REST-compatibel ID terug.
      const ewsId = (item as any).itemId || item.itemId;
      const itemId =
        typeof Office.context.mailbox.convertToRestId === "function"
          ? Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0)
          : ewsId;
      console.log("[save] itemId:", itemId, "sharedMailboxUser:", sharedMailboxUser);
      const emlBlob = await getMailMimeContent(itemId, sharedMailboxUser);

      // Auto-create Correspondentie submap en upload daarin
      const corrFolder = await getOrCreateSubFolder(
        siteId, driveId, selectedProjectId, AUTO_SUBFOLDER_NAME
      );

      await uploadToSharePoint(siteId, driveId, corrFolder.id, fileName, emlBlob);

      // Conversatie-history opslaan
      const conversationId = item.conversationId;
      if (conversationId) {
        const selectedProject = projects.find((p) => p.id === selectedProjectId);
        await saveConversationHistory({
          conversationId,
          normalizedSubject: normalizeSubject(item.subject || ""),
          workType,
          projectFolderId: selectedProjectId,
          projectFolderName: selectedProject?.name ?? "",
          archiveMailFolderId: selectedArchiveFolderId || undefined,
          archiveMailFolderName: mailFolders.find((f) => f.id === selectedArchiveFolderId)?.displayName,
          fileNameFields,
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
    siteId, driveId, selectedProjectId, selectedArchiveFolderId,
    fileName, workType, projects, mailFolders, sharedMailboxUser,
  ]);

  const canSave = siteId && driveId && selectedProjectId && saveStatus !== "saving";

  // --- Loading state ---
  if (orgLoading) {
    return (
      <div className={styles.root}>
        <Spinner label="Organisatie detecteren..." />
      </div>
    );
  }

  // --- Fout state ---
  if (orgError) {
    return (
      <div className={styles.root}>
        <div className={styles.header}>
          <Text className={styles.title}>Mail naar SharePoint</Text>
        </div>
        <div className={`${styles.statusContainer} ${styles.statusError}`}>
          <DismissCircle24Filled />
          <Text>{orgError}</Text>
        </div>
      </div>
    );
  }

  // --- Render ---
  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text className={styles.title}>Mail naar SharePoint</Text>
        {detectedOrg && <Text className={styles.orgBadge}>{detectedOrg.name}</Text>}
      </div>

      <div className={styles.form}>
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
              disabled={!detectedOrg?.projecten}
              title={!detectedOrg?.projecten ? "Nog niet beschikbaar" : undefined}
              style={{ flexGrow: 1 }}
            >
              Projecten
            </ToggleButton>
          </div>
        </div>

        {/* Fallback: bibliotheek handmatig kiezen als auto-detectie mislukt */}
        {libraries.length > 0 && !driveId && (
          <div className={styles.fieldGroup}>
            <Text className={styles.label}>Bibliotheek</Text>
            <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorPaletteYellowForeground1 }}>
              Bibliotheek niet automatisch gevonden — kies handmatig:
            </Text>
            <Dropdown
              placeholder="Selecteer bibliotheek..."
              value={libraries.find((l) => l.id === selectedLibraryId)?.name ?? ""}
              selectedOptions={selectedLibraryId ? [selectedLibraryId] : []}
              onOptionSelect={handleFallbackLibraryChange}
            >
              {libraries.map((lib) => (
                <Option key={lib.id} value={lib.id}>{lib.name}</Option>
              ))}
            </Dropdown>
          </div>
        )}

        {/* Project selectie */}
        <div className={styles.fieldGroup}>
          <Text className={styles.label}>Project</Text>
          {loadingProjects ? (
            <Spinner size="tiny" label="Projecten laden..." />
          ) : (
            <Combobox
              placeholder={
                projects.length === 0
                  ? "Geen projecten gevonden"
                  : "Typ om te filteren (bijv. 25-001)..."
              }
              value={projectInput}
              selectedOptions={selectedProjectId ? [selectedProjectId] : []}
              onOptionSelect={handleProjectSelect}
              onChange={handleProjectInputChange}
              disabled={projects.length === 0}
            >
              {filteredProjects.map((project) => (
                <Option key={project.id} value={project.id}>{project.name}</Option>
              ))}
            </Combobox>
          )}
          {selectedProjectId && (
            <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
              Opslaan in: {projects.find((p) => p.id === selectedProjectId)?.name} / {AUTO_SUBFOLDER_NAME}
            </Text>
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
