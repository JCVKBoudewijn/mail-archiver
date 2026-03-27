/** SharePoint site informatie */
export interface SharePointSite {
  id: string;
  displayName: string;
  webUrl: string;
}

/** Project folder in SharePoint */
export interface ProjectFolder {
  id: string;
  name: string;
  webUrl: string;
}

/** SharePoint documentbibliotheek */
export interface Library {
  id: string;
  name: string;
  webUrl: string;
}

/** Submap binnen een project */
export interface SubFolder {
  id: string;
  name: string;
  webUrl: string;
}

/** Outlook mail folder voor archivering */
export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
}

/** Opgeslagen instellingen per conversatie (roaming settings) */
export interface ConversationHistory {
  conversationId: string;
  normalizedSubject?: string;
  siteId: string;
  siteName: string;
  libraryId: string;
  libraryName: string;
  projectFolderId: string;
  projectFolderName: string;
  subFolderId: string;
  subFolderName: string;
  archiveMailFolderId?: string;
  archiveMailFolderName?: string;
  timestamp: number;
}

/** Beschikbare metadata velden voor bestandsnaam */
export type FileNameField = "date" | "subject" | "sender" | "recipient";

/** Configuratie voor bestandsnaam opbouw */
export interface FileNameConfig {
  fields: FileNameField[];  // volgorde + welke velden actief zijn
}

export const DEFAULT_FILENAME_CONFIG: FileNameConfig = {
  fields: ["date", "subject"],
};

/** Alle roaming settings data */
export interface RoamingData {
  conversations: ConversationHistory[];
  fileNameConfig?: FileNameConfig;
}

/** Status van de opslag-actie */
export type SaveStatus = "idle" | "saving" | "success" | "error";

/** Bibliotheek configuratie binnen een organisatie */
export interface OrgLibraryConfig {
  /** Naam van de SharePoint drive/bibliotheek */
  libraryName: string;
  /** Optioneel subpad binnen de bibliotheek, bijv. "2. Werken" */
  subPath?: string;
}

/** Organisatie configuratie voor auto-detectie */
export interface OrgConfig {
  /** Weergavenaam van de organisatie */
  name: string;
  /** E-maildomeinen die bij deze organisatie horen */
  emailDomains: string[];
  /** SharePoint site pad, bijv. "sites/Entropal717" */
  siteUrl: string;
  /** Bibliotheek configuratie voor Werken */
  werken: OrgLibraryConfig;
  /** Bibliotheek configuratie voor Projecten (null = nog niet beschikbaar) */
  projecten: OrgLibraryConfig | null;
}

/** Applicatie configuratie */
export interface AppConfig {
  /** SharePoint tenant hostname */
  tenantHostname: string;
  /** Organisatie configuraties voor auto-detectie */
  orgConfigs: OrgConfig[];
}
