/**
 * SSO Authenticatie service.
 * Gebruikt Office.js SSO om een access token te verkrijgen voor Microsoft Graph.
 * Fallback: PKCE authorization code flow via Office dialog.
 *
 * Tokens worden opgeslagen in localStorage zodat ze behouden blijven
 * wanneer de task pane opnieuw geopend wordt (sessionStorage gaat dan verloren).
 */

const CLIENT_ID = "06e23f21-f875-4425-aca3-ccd0b06bb24f";
const TENANT_ID = "eba9b46b-0bb0-493e-8724-854a60012ad4";
const TOKEN_ENDPOINT = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const AUTH_ENDPOINT = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;

const GRAPH_SCOPES = [
  "Files.ReadWrite.All",
  "Sites.Read.All",
  "Mail.ReadWrite",
  "Mail.ReadWrite.Shared",
];

// Alle keys in localStorage zodat ze overleven tussen task pane sessies
const TOKEN_KEY = "mailsp_token";
const TOKEN_EXPIRY_KEY = "mailsp_token_expiry";
const TOKEN_SOURCE_KEY = "mailsp_token_source";
const REFRESH_TOKEN_KEY = "mailsp_refresh_token";
const TOKEN_SCOPES_KEY = "mailsp_token_scopes";
const SSO_FAILED_KEY = "mailsp_sso_failed";

let cachedToken: string | null = localStorage.getItem(TOKEN_KEY);
let tokenExpiry: number = Number(localStorage.getItem(TOKEN_EXPIRY_KEY) || "0");

/** Herstelt ssoFailedForGraph uit localStorage */
let ssoFailedForGraph: boolean = localStorage.getItem(SSO_FAILED_KEY) === "true";

// ── Token opslag helpers ─────────────────────────────────────────────────────

function storeToken(token: string, source: string): void {
  cachedToken = token;
  tokenExpiry = Date.now() + 28_800_000; // 8 uur
  localStorage.setItem(TOKEN_KEY, token);
  localStorage.setItem(TOKEN_EXPIRY_KEY, tokenExpiry.toString());
  localStorage.setItem(TOKEN_SOURCE_KEY, source);
  localStorage.setItem(TOKEN_SCOPES_KEY, JSON.stringify(GRAPH_SCOPES));
}

/** Check of opgeslagen token voldoende scopes heeft; clear anders */
function validateTokenScopes(): void {
  const stored = localStorage.getItem(TOKEN_SCOPES_KEY);
  const storedScopes = stored ? JSON.parse(stored) : [];
  if (JSON.stringify(GRAPH_SCOPES.sort()) !== JSON.stringify(storedScopes.sort())) {
    console.log("[auth] Scopes gewijzigd — tokens gewist");
    clearTokenCache(true);
  }
}

// ── PKCE helpers ─────────────────────────────────────────────────────────────

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

// ── Public API ───────────────────────────────────────────────────────────────

/**
 * Verkrijg een access token.
 * Volgorde: cached → refresh token (stil) → SSO → PKCE dialog.
 */
export async function getAccessToken(): Promise<string> {
  validateTokenScopes();

  // Cached token nog geldig? (5 min marge)
  if (cachedToken && Date.now() < tokenExpiry - 300_000) {
    return cachedToken;
  }

  // Probeer silent refresh via opgeslagen refresh token (geen popup)
  const storedRefreshToken = localStorage.getItem(REFRESH_TOKEN_KEY);
  if (storedRefreshToken) {
    try {
      console.log("[auth] Silent refresh via refresh token");
      return await refreshAccessToken(storedRefreshToken);
    } catch (error) {
      console.warn("[auth] Refresh token mislukt:", error);
      localStorage.removeItem(REFRESH_TOKEN_KEY);
    }
  }

  // SSO eerder gefaald voor Graph? Sla over, direct PKCE.
  if (ssoFailedForGraph) {
    console.log("[auth] SSO eerder gefaald, direct PKCE");
    return await fallbackToDialogAuth();
  }

  // Probeer SSO
  try {
    const bootstrapToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });

    storeToken(bootstrapToken, "sso");
    return bootstrapToken;
  } catch (error: any) {
    const code = error?.code;

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

  const response = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!response.ok) {
    throw new Error(`Refresh token mislukt (${response.status})`);
  }

  const data = await response.json();
  storeToken(data.access_token, "refresh");

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
    `${AUTH_ENDPOINT}?` +
    `client_id=${encodeURIComponent(CLIENT_ID)}` +
    `&response_type=code` +
    `&scope=${encodeURIComponent(GRAPH_SCOPES.join(" ") + " offline_access")}` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&code_challenge=${challenge}` +
    `&code_challenge_method=S256` +
    `&state=${state}` +
    `&prompt=none`; // Probeer silent auth, geen login-scherm als sessie al bestaat

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

              const { accessToken, refreshToken } = await exchangeCodeForToken(message.code, verifier, redirectUri);
              storeToken(accessToken, "pkce");
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

  const response = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error_description || `Token uitwisseling mislukt (${response.status})`);
  }

  const data = await response.json();
  return { accessToken: data.access_token, refreshToken: data.refresh_token };
}

/** Reset de token cache */
export function clearTokenCache(includeRefreshToken = false): void {
  cachedToken = null;
  tokenExpiry = 0;
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(TOKEN_EXPIRY_KEY);
  localStorage.removeItem(TOKEN_SOURCE_KEY);
  localStorage.removeItem(TOKEN_SCOPES_KEY);
  if (includeRefreshToken) {
    localStorage.removeItem(REFRESH_TOKEN_KEY);
  }
}

/**
 * Markeer dat het SSO-token niet werkt voor Graph API (401).
 * Persistent in localStorage zodat het niet vergeten wordt bij heropen.
 */
export function markSsoFailedForGraph(): void {
  ssoFailedForGraph = true;
  localStorage.setItem(SSO_FAILED_KEY, "true");
}

/** Geeft aan of het huidige token via SSO is verkregen */
export function isTokenFromSso(): boolean {
  return localStorage.getItem(TOKEN_SOURCE_KEY) === "sso";
}
