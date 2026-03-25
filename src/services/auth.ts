/**
 * SSO Authenticatie service.
 * Gebruikt Office.js SSO om een access token te verkrijgen voor Microsoft Graph.
 */

const GRAPH_SCOPES = [
  "Files.ReadWrite.All",
  "Sites.Read.All",
  "Mail.ReadWrite",
];

let cachedToken: string | null = null;
let tokenExpiry: number = 0;

/**
 * Verkrijg een access token via Office SSO.
 * Probeert eerst SSO; bij falen wordt een consent dialog getoond.
 */
export async function getAccessToken(): Promise<string> {
  // Return cached token als deze nog geldig is (5 min marge)
  if (cachedToken && Date.now() < tokenExpiry - 300_000) {
    return cachedToken;
  }

  try {
    // Stap 1: Verkrijg SSO bootstrap token van Office
    const bootstrapToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });

    // Het SSO token kan direct gebruikt worden voor Graph API calls
    // als de admin consent heeft gegeven voor de benodigde scopes.
    cachedToken = bootstrapToken;
    // SSO tokens zijn typisch 1 uur geldig
    tokenExpiry = Date.now() + 3_600_000;

    return bootstrapToken;
  } catch (error: any) {
    const code = error?.code;

    // 13001: Gebruiker niet ingelogd
    // 13002: Consent nodig
    // 13003: Niet-ondersteunde omgeving
    if (code === 13001 || code === 13002 || code === 13003) {
      return await fallbackToDialogAuth();
    }

    console.error("SSO auth failed:", error);
    throw new Error(
      `Authenticatie mislukt (code: ${code}). Probeer de add-in opnieuw te openen.`
    );
  }
}

/**
 * Fallback: gebruik MSAL dialog voor consent/login.
 * Dit opent een popup waar de gebruiker kan inloggen.
 */
async function fallbackToDialogAuth(): Promise<string> {
  return new Promise((resolve, reject) => {
    const dialogUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
      `client_id=${encodeURIComponent("06e23f21-f875-4425-aca3-ccd0b06bb24f")}` +
      `&response_type=token` +
      `&scope=${encodeURIComponent(GRAPH_SCOPES.join(" "))}` +
      `&redirect_uri=${encodeURIComponent(window.location.origin + "/taskpane.html")}` +
      `&prompt=consent`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 40 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error("Kan login dialog niet openen."));
          return;
        }

        const dialog = result.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: any) => {
            dialog.close();
            try {
              const message = JSON.parse(arg.message);
              if (message.accessToken) {
                cachedToken = message.accessToken;
                tokenExpiry = Date.now() + 3_600_000;
                resolve(message.accessToken);
              } else {
                reject(new Error("Geen access token ontvangen."));
              }
            } catch {
              reject(new Error("Ongeldig antwoord van login dialog."));
            }
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          () => {
            reject(new Error("Login dialog is gesloten."));
          }
        );
      }
    );
  });
}

/** Reset de token cache (bijv. bij logout of token error) */
export function clearTokenCache(): void {
  cachedToken = null;
  tokenExpiry = 0;
}
