import { AppConfig } from "./types";

/**
 * Applicatieconfiguratie.
 * Pas deze waarden aan voor jouw organisatie.
 */
export const APP_CONFIG: AppConfig = {
  // Standaard SharePoint hostname voor J.C. van Kessel
  defaultSiteHostname: "jcvankessel.sharepoint.com",

  // Configuratie per site - basispad waar projectmappen staan
  siteConfigs: [
    {
      siteId: "", // Wordt dynamisch opgehaald
      siteName: "Hoofdsite",
      basePath: "Documenten/2. Werken",
    },
  ],

  // Fallback basispad als site niet specifiek geconfigureerd is
  defaultBasePath: "Documenten/2. Werken",
};

/** Microsoft Graph API base URL */
export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

/** Maximale grootte van roaming settings in bytes (Office limiet) */
export const ROAMING_SETTINGS_MAX_BYTES = 32_768; // 32 KB

/** Projectnummer regex: YY-XXX formaat */
export const PROJECT_NUMBER_REGEX = /^\d{2}-\d{3}/;
