/**
 * RoamingSettings Manager.
 * Slaat per ConversationID de laatst gekozen site, project en map op.
 * Implementeert FIFO: oudste entries worden verwijderd als de 32KB limiet bereikt wordt.
 */

import { ROAMING_SETTINGS_MAX_BYTES } from "../config";
import type { ConversationHistory, FileNameConfig, RoamingData } from "../types";
import { DEFAULT_FILENAME_CONFIG } from "../types";

const SETTINGS_KEY = "mailToSharePoint";

/** Schat de byte-grootte van een string (UTF-8) */
function estimateByteSize(str: string): number {
  return new Blob([str]).size;
}

/** Lees alle opgeslagen conversatie-histories */
export function loadConversationHistory(): ConversationHistory[] {
  const settings = Office.context.roamingSettings;
  const data = settings.get(SETTINGS_KEY) as RoamingData | undefined;
  return data?.conversations ?? [];
}

/**
 * Sla een conversatie-history op.
 * Bij een bestaand conversationId wordt de entry bijgewerkt.
 * Bij een nieuw conversationId wordt deze toegevoegd.
 * FIFO: oudste entries worden verwijderd als de limiet bereikt wordt.
 */
export async function saveConversationHistory(
  entry: ConversationHistory
): Promise<void> {
  const settings = Office.context.roamingSettings;
  let conversations = loadConversationHistory();

  // Update bestaande entry of voeg nieuwe toe
  const existingIndex = conversations.findIndex(
    (c) => c.conversationId === entry.conversationId
  );

  if (existingIndex >= 0) {
    conversations[existingIndex] = { ...entry, timestamp: Date.now() };
  } else {
    conversations.push({ ...entry, timestamp: Date.now() });
  }

  // Sorteer op timestamp (nieuwste eerst) voor FIFO verwijdering
  conversations.sort((a, b) => b.timestamp - a.timestamp);

  // FIFO: verwijder oudste entries totdat we onder de limiet zijn
  let data: RoamingData = { conversations };
  let serialized = JSON.stringify(data);

  while (
    estimateByteSize(serialized) > ROAMING_SETTINGS_MAX_BYTES * 0.9 &&
    conversations.length > 1
  ) {
    // Verwijder de oudste entry (laatste in de gesorteerde array)
    conversations.pop();
    data = { conversations };
    serialized = JSON.stringify(data);
  }

  // Sla op in roaming settings
  settings.set(SETTINGS_KEY, data);

  // Persist naar de server
  return new Promise((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error("Kan roaming settings niet opslaan."));
      }
    });
  });
}

/**
 * Zoek de opgeslagen history voor een specifiek conversatie-ID.
 * Returns undefined als er geen history is.
 */
export function getHistoryForConversation(
  conversationId: string
): ConversationHistory | undefined {
  const conversations = loadConversationHistory();
  return conversations.find((c) => c.conversationId === conversationId);
}

/** Lees de opgeslagen bestandsnaam configuratie */
export function loadFileNameConfig(): FileNameConfig {
  const settings = Office.context.roamingSettings;
  const data = settings.get(SETTINGS_KEY) as RoamingData | undefined;
  return data?.fileNameConfig ?? DEFAULT_FILENAME_CONFIG;
}

/** Sla de bestandsnaam configuratie op */
export async function saveFileNameConfig(config: FileNameConfig): Promise<void> {
  const settings = Office.context.roamingSettings;
  const data = (settings.get(SETTINGS_KEY) as RoamingData) ?? { conversations: [] };
  settings.set(SETTINGS_KEY, { ...data, fileNameConfig: config });

  return new Promise((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error("Kan bestandsnaam configuratie niet opslaan."));
      }
    });
  });
}

/** Verwijder alle opgeslagen conversation histories */
export async function clearAllHistory(): Promise<void> {
  const settings = Office.context.roamingSettings;
  settings.remove(SETTINGS_KEY);

  return new Promise((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error("Kan roaming settings niet wissen."));
      }
    });
  });
}
