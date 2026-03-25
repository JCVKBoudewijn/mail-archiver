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

/** Alle roaming settings data */
export interface RoamingData {
  conversations: ConversationHistory[];
}

/** Status van de opslag-actie */
export type SaveStatus = "idle" | "saving" | "success" | "error";

/** Configuratie per SharePoint site */
export interface SiteConfig {
  siteId: string;
  siteName: string;
  /** Basispad voor projectmappen, bijv. "Documenten/2. Werken" */
  basePath: string;
}

/** Applicatie configuratie */
export interface AppConfig {
  /** Standaard SharePoint site hostname, bijv. "jcvankessel.sharepoint.com" */
  defaultSiteHostname: string;
  /** Configuratie per site */
  siteConfigs: SiteConfig[];
  /** Standaard basispad als site niet in configs staat */
  defaultBasePath: string;
}
