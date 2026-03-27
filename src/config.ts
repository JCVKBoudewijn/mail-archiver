import { AppConfig } from "./types";

/**
 * Applicatieconfiguratie.
 * Organisatie-detectie op basis van e-maildomein.
 */
export const APP_CONFIG: AppConfig = {
  tenantHostname: "jcvankessel.sharepoint.com",

  orgConfigs: [
    {
      name: "Entropal",
      emailDomains: ["entropal.nl"],
      siteUrl: "sites/Entropal717",
      werken: {
        libraryName: "Gedeelde documenten",
        subPath: "2. Werken",
      },
      projecten: {
        libraryName: "Gedeelde documenten",
        subPath: "1. Projecten",
      },
    },
    {
      name: "SolarComfort",
      emailDomains: ["solarcomfort.nl"],
      siteUrl: "sites/SolarComfortSolarTekPlanning",
      werken: {
        libraryName: "SC  Werken",
      },
      projecten: null, // nog niet beschikbaar
    },
    {
      name: "SolarTek",
      emailDomains: ["solartek.nl"],
      siteUrl: "sites/SolarComfortSolarTekPlanning",
      werken: {
        libraryName: "Projecten",
      },
      projecten: null, // nog niet beschikbaar
    },
  ],
};

/** Microsoft Graph API base URL */
export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

/** Maximale grootte van roaming settings in bytes (Office limiet) */
export const ROAMING_SETTINGS_MAX_BYTES = 32_768; // 32 KB

/** Projectnummer regex: YY-XXX formaat */
export const PROJECT_NUMBER_REGEX = /^\d{2}-\d{3}/;

/** Naam van de automatisch aan te maken submap */
export const AUTO_SUBFOLDER_NAME = "Correspondentie";
