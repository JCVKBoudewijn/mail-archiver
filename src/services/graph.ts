/**
 * Microsoft Graph API service.
 * Alle communicatie met Graph API voor SharePoint en Mail operaties.
 */

import { getAccessToken, clearTokenCache, markSsoFailedForGraph, isTokenFromSso } from "./auth";
import { GRAPH_BASE_URL } from "../config";
import type {
  SharePointSite,
  Library,
  ProjectFolder,
  SubFolder,
  MailFolder,
  FileNameField,
} from "../types";

/** Generieke Graph API call met automatische token refresh */
async function graphFetch(
  endpoint: string,
  options: RequestInit = {}
): Promise<any> {
  const token = await getAccessToken();

  const response = await fetch(`${GRAPH_BASE_URL}${endpoint}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });

  if (response.status === 401) {
    // Als het token via SSO was verkregen, is het waarschijnlijk een bootstrap
    // token dat niet werkt voor Graph (geen server-side OBO exchange).
    // Markeer SSO als gefaald zodat volgende calls direct PKCE gebruiken.
    if (isTokenFromSso()) {
      console.warn("[graphFetch] SSO-token gaf 401 op Graph – schakel over naar PKCE");
      markSsoFailedForGraph();
    }

    clearTokenCache();
    const newToken = await getAccessToken();
    const retryResponse = await fetch(`${GRAPH_BASE_URL}${endpoint}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${newToken}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });

    if (!retryResponse.ok) {
      throw new Error(`Graph API fout: ${retryResponse.status} ${retryResponse.statusText}`);
    }
    return retryResponse.status === 204 ? null : retryResponse.json();
  }

  if (!response.ok) {
    const errorBody = await response.text();
    console.error(`[graphFetch] ${response.status} ${endpoint}`, errorBody);
    throw new Error(`Graph API fout: ${response.status} - ${errorBody}`);
  }

  return response.status === 204 ? null : response.json();
}

// ============================================================
// Gebruiker
// ============================================================

/** Haal het e-mailadres van de ingelogde gebruiker op */
export async function getUserEmail(): Promise<string> {
  const data = await graphFetch(`/me?$select=mail,userPrincipalName`);
  return data.mail || data.userPrincipalName || "";
}

// ============================================================
// SharePoint Sites
// ============================================================

/** Zoek SharePoint sites binnen de organisatie */
export async function searchSites(query: string): Promise<SharePointSite[]> {
  // Voor wildcard-query ("*"): gebruik lege search voor alle sites
  const searchQuery = query === "*" ? "" : query;
  const data = await graphFetch(
    `/sites?search=${encodeURIComponent(searchQuery)}&$select=id,displayName,webUrl&$top=100`
  );
  return (data.value || []).map((site: any) => ({
    id: site.id,
    displayName: site.displayName,
    webUrl: site.webUrl,
  }));
}

/** Haal een specifieke site op via hostname en pad */
export async function getSiteByHostname(
  hostname: string,
  path: string = ""
): Promise<SharePointSite> {
  const siteUrl = path
    ? `/sites/${hostname}:/${path}`
    : `/sites/${hostname}`;
  const data = await graphFetch(`${siteUrl}?$select=id,displayName,webUrl`);
  return {
    id: data.id,
    displayName: data.displayName,
    webUrl: data.webUrl,
  };
}

// ============================================================
// SharePoint Bibliotheken
// ============================================================

/** Haal mappen op in de root van een bibliotheek (voor subpad-kiezer) */
export async function getRootFolders(siteId: string, libraryId: string): Promise<SubFolder[]> {
  const data = await graphFetch(
    `/sites/${siteId}/drives/${libraryId}/root/children?$select=id,name,webUrl,folder&$top=100`
  );
  return (data.value || [])
    .filter((item: any) => item.folder !== undefined)
    .map((item: any) => ({ id: item.id, name: item.name, webUrl: item.webUrl }))
    .sort((a: SubFolder, b: SubFolder) => a.name.localeCompare(b.name));
}

/** Haal alle documentbibliotheken op van een SharePoint site */
export async function getLibraries(siteId: string): Promise<Library[]> {
  const data = await graphFetch(
    `/sites/${siteId}/drives?$select=id,name,webUrl`
  );
  return (data.value || [])
    .filter((d: any) => d.driveType === "documentLibrary" || d.name)
    .map((d: any) => ({
      id: d.id,
      name: d.name,
      webUrl: d.webUrl,
    }))
    .sort((a: Library, b: Library) => a.name.localeCompare(b.name));
}

// ============================================================
// SharePoint Folders (Projecten & Submappen)
// ============================================================

/**
 * Haal projectmappen op vanuit een bibliotheek, optioneel binnen een subpad.
 * libraryId: drive ID van de bibliotheek
 * subPath: optioneel subpad binnen de bibliotheek, bijv. "2. Werken"
 */
export async function getProjectFolders(
  siteId: string,
  libraryId: string,
  subPath?: string
): Promise<ProjectFolder[]> {
  const endpoint = subPath
    ? `/sites/${siteId}/drives/${libraryId}/root:/${encodeURIComponent(subPath)}:/children?$select=id,name,webUrl,folder&$top=500`
    : `/sites/${siteId}/drives/${libraryId}/root/children?$select=id,name,webUrl,folder&$top=500`;

  const data = await graphFetch(endpoint);

  return (data.value || [])
    .filter((item: any) => item.folder !== undefined)
    .map((item: any) => ({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
    }))
    .sort((a: ProjectFolder, b: ProjectFolder) => b.name.localeCompare(a.name));
}

/** Haal submappen op binnen een projectfolder */
export async function getSubFolders(
  siteId: string,
  folderId: string
): Promise<SubFolder[]> {
  const data = await graphFetch(
    `/sites/${siteId}/drive/items/${folderId}/children?$select=id,name,webUrl,folder&$top=100`
  );

  return (data.value || [])
    .filter((item: any) => item.folder !== undefined)
    .map((item: any) => ({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
    }))
    .sort((a: SubFolder, b: SubFolder) => a.name.localeCompare(b.name));
}

/** Maak een submap aan binnen een folder (als die nog niet bestaat) */
export async function getOrCreateSubFolder(
  siteId: string,
  driveId: string,
  parentFolderId: string,
  folderName: string
): Promise<SubFolder> {
  // Eerst kijken of de map al bestaat
  const existing = await getSubFoldersInDrive(siteId, driveId, parentFolderId);
  const found = existing.find((f) => f.name.toLowerCase() === folderName.toLowerCase());
  if (found) return found;

  // Map aanmaken
  const data = await graphFetch(
    `/sites/${siteId}/drives/${driveId}/items/${parentFolderId}/children`,
    {
      method: "POST",
      body: JSON.stringify({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      }),
    }
  );

  return { id: data.id, name: data.name, webUrl: data.webUrl };
}

/** Haal submappen op binnen een folder via drive ID */
export async function getSubFoldersInDrive(
  siteId: string,
  driveId: string,
  folderId: string
): Promise<SubFolder[]> {
  const data = await graphFetch(
    `/sites/${siteId}/drives/${driveId}/items/${folderId}/children?$select=id,name,webUrl,folder&$top=100`
  );

  return (data.value || [])
    .filter((item: any) => item.folder !== undefined)
    .map((item: any) => ({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
    }))
    .sort((a: SubFolder, b: SubFolder) => a.name.localeCompare(b.name));
}

/** Haal de drive ID op van een bibliotheek via naam */
export async function getLibraryByName(
  siteId: string,
  libraryName: string
): Promise<Library | null> {
  const libs = await getLibraries(siteId);
  return libs.find((l) => l.name === libraryName) || null;
}

// ============================================================
// E-mail ophalen en uploaden naar SharePoint
// ============================================================

/** Haal de MIME content (.eml) van een e-mail op */
export async function getMailMimeContent(messageId: string, mailboxUser?: string): Promise<Blob> {
  const token = await getAccessToken();
  const mailboxPath = mailboxUser ? `/users/${encodeURIComponent(mailboxUser)}` : "/me";
  const response = await fetch(
    `${GRAPH_BASE_URL}${mailboxPath}/messages/${messageId}/$value`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  if (!response.ok) {
    throw new Error(`Kan e-mail niet ophalen: ${response.status}`);
  }

  return response.blob();
}

/**
 * Upload een bestand naar een specifieke SharePoint folder.
 * Gebruikt een upload session voor bestanden > 4MB.
 */
export async function uploadToSharePoint(
  siteId: string,
  folderId: string,
  fileName: string,
  content: Blob
): Promise<string> {
  const sanitizedName = sanitizeFileName(fileName);

  if (content.size < 4_000_000) {
    // Kleine bestanden: directe upload
    const token = await getAccessToken();
    const response = await fetch(
      `${GRAPH_BASE_URL}/sites/${siteId}/drive/items/${folderId}:/${encodeURIComponent(sanitizedName)}:/content`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/octet-stream",
        },
        body: content,
      }
    );

    if (!response.ok) {
      throw new Error(`Upload mislukt: ${response.status}`);
    }

    const data = await response.json();
    return data.webUrl;
  }

  // Grote bestanden: upload session
  const sessionData = await graphFetch(
    `/sites/${siteId}/drive/items/${folderId}:/${encodeURIComponent(sanitizedName)}:/createUploadSession`,
    {
      method: "POST",
      body: JSON.stringify({
        item: { "@microsoft.graph.conflictBehavior": "rename" },
      }),
    }
  );

  const uploadUrl = sessionData.uploadUrl;
  const token = await getAccessToken();
  const arrayBuffer = await content.arrayBuffer();
  const chunkSize = 3_276_800; // ~3.1 MB chunks
  let offset = 0;

  while (offset < arrayBuffer.byteLength) {
    const end = Math.min(offset + chunkSize, arrayBuffer.byteLength);
    const chunk = arrayBuffer.slice(offset, end);

    const uploadResponse = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Length": chunk.byteLength.toString(),
        "Content-Range": `bytes ${offset}-${end - 1}/${arrayBuffer.byteLength}`,
      },
      body: chunk,
    });

    if (!uploadResponse.ok && uploadResponse.status !== 202) {
      throw new Error(`Upload chunk mislukt bij offset ${offset}`);
    }

    offset = end;
  }

  return sanitizedName;
}

// ============================================================
// Outlook Mail Folders (voor archivering)
// ============================================================

/** Haal alle mailmappen op (voor de archief-selector) */
export async function getMailFolders(mailboxUser?: string): Promise<MailFolder[]> {
  const mailboxPath = mailboxUser ? `/users/${encodeURIComponent(mailboxUser)}` : "/me";
  const data = await graphFetch(
    `${mailboxPath}/mailFolders?$select=id,displayName,parentFolderId&$top=100`
  );

  return (data.value || []).map((folder: any) => ({
    id: folder.id,
    displayName: folder.displayName,
    parentFolderId: folder.parentFolderId,
  }));
}

/** Verplaats een e-mail naar een andere map */
export async function moveMailToFolder(
  messageId: string,
  destinationFolderId: string,
  mailboxUser?: string
): Promise<void> {
  const mailboxPath = mailboxUser ? `/users/${encodeURIComponent(mailboxUser)}` : "/me";
  await graphFetch(`${mailboxPath}/messages/${messageId}/move`, {
    method: "POST",
    body: JSON.stringify({ destinationId: destinationFolderId }),
  });
}

// ============================================================
// Helpers
// ============================================================

/** Verwijder ongeldige tekens uit bestandsnamen */
function sanitizeFileName(name: string): string {
  return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").substring(0, 250);
}

/**
 * Genereer de bestandsnaam voor een e-mail op basis van geconfigureerde velden.
 * Velden worden samengevoegd met " - " als separator.
 */
export function generateEmailFileName(
  subject: string,
  dateReceived: Date,
  fields: FileNameField[] = ["date", "subject"],
  sender: string = "",
  recipient: string = ""
): string {
  const pad = (n: number) => n.toString().padStart(2, "0");
  const y = dateReceived.getFullYear();
  const mo = pad(dateReceived.getMonth() + 1);
  const d = pad(dateReceived.getDate());
  const hh = pad(dateReceived.getHours());
  const mm = pad(dateReceived.getMinutes());

  const clean = (s: string) => s.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").trim();

  const parts = fields.map((field) => {
    switch (field) {
      case "date":      return `${y}-${mo}-${d}-${hh}${mm}`;
      case "subject":   return clean(subject);
      case "sender":    return clean(sender);
      case "recipient": return clean(recipient);
    }
  }).filter(Boolean);

  return `${parts.join(" - ")}.eml`;
}
