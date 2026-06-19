/**
 * Salesforce OAuth 2.0 (PKCE, public client) for the Office task pane.
 *
 * Flow:
 *  1. Task pane opens the org's /services/oauth2/authorize in an Office dialog
 *     (Office.context.ui.displayDialogAsync) with a PKCE challenge.
 *  2. Salesforce redirects to our hosted auth-callback.html, which posts the
 *     ?code & state back via Office.context.ui.messageParent.
 *  3. The task pane exchanges the code (+ verifier) for tokens by calling the
 *     org's /services/oauth2/token directly from the browser. This requires the
 *     org to enable "CORS for OAuth Endpoints" (with office.sliick.com on the
 *     CORS allowlist). Running the exchange here means it originates from the
 *     user's own IP, so IP-enforcing apps accept it.
 *
 * Access token lives in sessionStorage (cleared when Office closes the pane);
 * refresh token in localStorage so reopening Word doesn't force a re-login.
 */

import {
  buildAuthorizeUrl,
  computeCodeChallenge,
  generateCodeVerifier,
  generateState,
} from "./pkce";

export interface TokenSet {
  accessToken: string;
  refreshToken?: string;
  instanceUrl: string;
  issuedAt: number;
}

interface DialogResultMessage {
  code?: string;
  state?: string;
  error?: string;
  errorDescription?: string;
}

const ACCESS_KEY = "sliick.auth.access.v1";
const REFRESH_KEY = "sliick.auth.refresh.v1";

export function getStoredTokens(): TokenSet | null {
  try {
    const raw = sessionStorage.getItem(ACCESS_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as TokenSet;
  } catch {
    return null;
  }
}

function storeTokens(tokens: TokenSet): void {
  sessionStorage.setItem(ACCESS_KEY, JSON.stringify(tokens));
  if (tokens.refreshToken) {
    localStorage.setItem(REFRESH_KEY, tokens.refreshToken);
  }
}

export function clearTokens(): void {
  sessionStorage.removeItem(ACCESS_KEY);
  localStorage.removeItem(REFRESH_KEY);
}

function redirectUri(): string {
  return `${window.location.origin}/auth-callback.html`;
}

/** Interactive login via Office dialog. Resolves with a usable token set. */
export async function login(orgUrl: string, clientId: string): Promise<TokenSet> {
  const verifier = generateCodeVerifier();
  const challenge = await computeCodeChallenge(verifier);
  const state = generateState();

  const authorizeUrl = buildAuthorizeUrl({
    orgUrl,
    clientId,
    redirectUri: redirectUri(),
    codeChallenge: challenge,
    state,
  });

  // Office's dialog only permits opening a URL on our own domain. The per-org
  // Salesforce authorize domain can't be enumerated in AppDomains, so we open a
  // same-origin bootstrap page that redirects to the authorize URL (see
  // auth-start.html / start.ts). AppDomains governs only the initial dialog URL.
  const dialogUrl = `${window.location.origin}/auth-start?u=${encodeURIComponent(authorizeUrl)}`;

  const message = await openAuthDialog(dialogUrl);
  if (message.error) {
    throw new Error(
      `Salesforce login failed: ${message.error} ${message.errorDescription ?? ""}`.trim(),
    );
  }
  if (!message.code || message.state !== state) {
    throw new Error("Salesforce login failed: missing code or state mismatch.");
  }
  return exchangeCode(orgUrl, clientId, message.code, verifier);
}

function openAuthDialog(url: string): Promise<DialogResultMessage> {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { height: 70, width: 40, promptBeforeOpen: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(`Could not open login dialog: ${result.error?.message}`));
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: { message: string } | { error: number }) => {
            dialog.close();
            if ("message" in arg) {
              try {
                resolve(JSON.parse(arg.message) as DialogResultMessage);
              } catch {
                reject(new Error("Malformed message from login dialog."));
              }
            } else {
              reject(new Error(`Login dialog error code ${arg.error}`));
            }
          },
        );
        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (arg: { message: string; origin: string | undefined } | { error: number }) => {
            const code = "error" in arg ? arg.error : undefined;
            // 12006 = user closed the dialog
            reject(
              new Error(
                code === 12006
                  ? "Login cancelled."
                  : `Login dialog closed unexpectedly (code ${code ?? "unknown"}).`,
              ),
            );
          },
        );
      },
    );
  });
}

async function exchangeCode(
  orgUrl: string,
  clientId: string,
  code: string,
  verifier: string,
): Promise<TokenSet> {
  const body = new URLSearchParams({
    grant_type: "authorization_code",
    code,
    client_id: clientId,
    redirect_uri: redirectUri(),
    code_verifier: verifier,
  });
  const tokens = await tokenRequest(orgUrl, body);
  storeTokens(tokens);
  return tokens;
}

/** Refresh the access token; throws if no refresh token is stored. */
export async function refreshAccessToken(
  orgUrl: string,
  clientId: string,
): Promise<TokenSet> {
  const refreshToken = localStorage.getItem(REFRESH_KEY);
  if (!refreshToken) throw new Error("No refresh token; interactive login required.");
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    client_id: clientId,
  });
  const tokens = await tokenRequest(orgUrl, body);
  // Salesforce may not return a new refresh token on refresh; keep the old one.
  if (!tokens.refreshToken) tokens.refreshToken = refreshToken;
  storeTokens(tokens);
  return tokens;
}

async function tokenRequest(orgUrl: string, body: URLSearchParams): Promise<TokenSet> {
  // Direct browser call to the org's token endpoint. Requires the org to have
  // "Enable CORS for OAuth Endpoints" on with office.sliick.com allowlisted.
  // Running it here (not server-side) means the exchange comes from the user's
  // own IP, so IP-enforcing apps accept it.
  const resp = await fetch(`${orgUrl}/services/oauth2/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  const json = (await resp.json()) as Record<string, string>;
  if (!resp.ok) {
    throw new Error(
      `Token request failed: ${json.error ?? resp.status} ${json.error_description ?? ""}`.trim(),
    );
  }
  if (!json.access_token || !json.instance_url) {
    throw new Error("Token response missing access_token/instance_url.");
  }
  return {
    accessToken: json.access_token,
    refreshToken: json.refresh_token,
    instanceUrl: json.instance_url,
    issuedAt: Number(json.issued_at ?? "0") || 0,
  };
}
