/**
 * SSO Authenticatie service.
 * Gebruikt Office.js SSO om een access token te verkrijgen voor Microsoft Graph.
 * Fallback: PKCE authorization code flow via Office dialog.
 */

const CLIENT_ID = "06e23f21-f875-4425-aca3-ccd0b06bb24f";

const GRAPH_SCOPES = [
  "Files.ReadWrite.All",
  "Sites.Read.All",
  "Mail.ReadWrite",
];

const TOKEN_KEY = "mailsp_token";
const TOKEN_EXPIRY_KEY = "mailsp_token_expiry";
const TOKEN_SOURCE_KEY = "mailsp_token_source";
const REFRESH_TOKEN_KEY = "mailsp_refresh_token";

let cachedToken: string | null = sessionStorage.getItem(TOKEN_KEY);
let tokenExpiry: number = Number(sessionStorage.getItem(TOKEN_EXPIRY_KEY) || "0");

/** Houdt bij of het SSO-token al gefaald heeft voor Graph (401).
 *  Als dat zo is, slaan we SSO over en gaan direct naar PKCE. */
let ssoFailedForGraph: boolean = false;

// ── PKCE helpers ──────────────────────────────────────────────────────────────

function base64UrlEncode(buffer: ArrayBuffer): string {
  return btoa(String.fromCharCode(...new Uint8Array(buffer)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

async function generateCodeVerifier(): Promise<string> {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return base64UrlEncode(array.buffer);
}

async function generateCodeChallenge(verifier: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return base64UrlEncode(digest);
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Verkrijg een access token via Office SSO.
 * Probeert eerst SSO; bij falen refresh token; daarna PKCE dialog flow.
 */
export async function getAccessToken(): Promise<string> {
  if (cachedToken && Date.now() < tokenExpiry - 300_000) {
    return cachedToken;
  }

  // Probeer silent refresh via opgeslagen refresh token (geen popup)
  const storedRefreshToken = localStorage.getItem(REFRESH_TOKEN_KEY);
  if (storedRefreshToken) {
    try {
      return await refreshAccessToken(storedRefreshToken);
    } catch {
      localStorage.removeItem(REFRESH_TOKEN_KEY);
    }
  }

  // Als SSO-token eerder 401 gaf op Graph, sla SSO over en gebruik direct PKCE
  if (ssoFailedForGraph) {
    console.log("[auth] SSO eerder gefaald voor Graph, gebruik PKCE fallback");
    return await fallbackToDialogAuth();
  }

  try {
    const bootstrapToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });

    cachedToken = bootstrapToken;
    tokenExpiry = Date.now() + 3_600_000;
    sessionStorage.setItem(TOKEN_KEY, bootstrapToken);
    sessionStorage.setItem(TOKEN_EXPIRY_KEY, tokenExpiry.toString());
    sessionStorage.setItem(TOKEN_SOURCE_KEY, "sso");
    return bootstrapToken;
  } catch (error: any) {
    const code = error?.code;

    // 13001: Gebruiker niet ingelogd
    // 13002: Consent nodig
    // 13003: Niet-ondersteunde omgeving
    // 13006: SSO mislukt (geen on-premises AD / intranet zone)
    // 13012: SSO niet ondersteund in deze context
    if ([13001, 13002, 13003, 13006, 13012].includes(code)) {
      return await fallbackToDialogAuth();
    }

    console.error("SSO auth failed:", error);
    throw new Error(
      `Authenticatie mislukt (code: ${code}). Probeer de add-in opnieuw te openen.`
    );
  }
}

async function refreshAccessToken(refreshToken: string): Promise<string> {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: GRAPH_SCOPES.join(" ") + " offline_access",
  });

  const response = await fetch(
    "https://login.microsoftonline.com/eba9b46b-0bb0-493e-8724-854a60012ad4/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  if (!response.ok) {
    throw new Error(`Refresh token mislukt (${response.status})`);
  }

  const data = await response.json();

  cachedToken = data.access_token;
  tokenExpiry = Date.now() + 3_600_000;
  sessionStorage.setItem(TOKEN_KEY, data.access_token);
  sessionStorage.setItem(TOKEN_EXPIRY_KEY, tokenExpiry.toString());
  sessionStorage.setItem(TOKEN_SOURCE_KEY, "refresh");

  if (data.refresh_token) {
    localStorage.setItem(REFRESH_TOKEN_KEY, data.refresh_token);
  }

  return data.access_token;
}

/**
 * Fallback: PKCE authorization code flow via Office dialog.
 */
async function fallbackToDialogAuth(): Promise<string> {
  const verifier = await generateCodeVerifier();
  const challenge = await generateCodeChallenge(verifier);
  const state = base64UrlEncode(crypto.getRandomValues(new Uint8Array(16)).buffer);
  const redirectUri = window.location.origin + "/auth-callback.html";

  const dialogUrl =
    `https://login.microsoftonline.com/eba9b46b-0bb0-493e-8724-854a60012ad4/oauth2/v2.0/authorize?` +
    `client_id=${encodeURIComponent(CLIENT_ID)}` +
    `&response_type=code` +
    `&scope=${encodeURIComponent(GRAPH_SCOPES.join(" ") + " offline_access")}` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&code_challenge=${challenge}` +
    `&code_challenge_method=S256` +
    `&state=${state}`;

  return new Promise((resolve, reject) => {
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
          async (arg: any) => {
            dialog.close();
            try {
              const message = JSON.parse(arg.message);

              if (message.error) {
                reject(new Error(`Login mislukt: ${message.errorDescription || message.error}`));
                return;
              }

              if (!message.code) {
                reject(new Error("Geen authorization code ontvangen."));
                return;
              }

              // Wissel de code in voor een access token
              const { accessToken, refreshToken } = await exchangeCodeForToken(message.code, verifier, redirectUri);
              cachedToken = accessToken;
              tokenExpiry = Date.now() + 3_600_000;
              sessionStorage.setItem(TOKEN_KEY, accessToken);
              sessionStorage.setItem(TOKEN_EXPIRY_KEY, tokenExpiry.toString());
              sessionStorage.setItem(TOKEN_SOURCE_KEY, "pkce");
              if (refreshToken) {
                localStorage.setItem(REFRESH_TOKEN_KEY, refreshToken);
              }
              resolve(accessToken);
            } catch (e: any) {
              reject(new Error(e.message || "Ongeldig antwoord van login dialog."));
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

async function exchangeCodeForToken(
  code: string,
  verifier: string,
  redirectUri: string
): Promise<{ accessToken: string; refreshToken?: string }> {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: verifier,
  });

  const response = await fetch(
    "https://login.microsoftonline.com/eba9b46b-0bb0-493e-8724-854a60012ad4/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error_description || `Token uitwisseling mislukt (${response.status})`);
  }

  const data = await response.json();
  return { accessToken: data.access_token, refreshToken: data.refresh_token };
}

/** Reset de token cache (bijv. bij logout of token error) */
export function clearTokenCache(includeRefreshToken = false): void {
  cachedToken = null;
  tokenExpiry = 0;
  sessionStorage.removeItem(TOKEN_KEY);
  sessionStorage.removeItem(TOKEN_EXPIRY_KEY);
  sessionStorage.removeItem(TOKEN_SOURCE_KEY);
  if (includeRefreshToken) {
    localStorage.removeItem(REFRESH_TOKEN_KEY);
  }
}

/**
 * Markeer dat het SSO-token niet werkt voor Graph API (401).
 * Volgende getAccessToken() calls zullen SSO overslaan en direct PKCE gebruiken.
 */
export function markSsoFailedForGraph(): void {
  ssoFailedForGraph = true;
}

/** Geeft aan of het huidige token via SSO is verkregen */
export function isTokenFromSso(): boolean {
  return sessionStorage.getItem(TOKEN_SOURCE_KEY) === "sso";
}
