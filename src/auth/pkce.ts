/**
 * OAuth 2.0 PKCE helpers (RFC 7636), S256 method.
 * Pure functions over WebCrypto — unit-testable outside Office.
 */

const VERIFIER_CHARS =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";

/** RFC 7636 §4.1: 43–128 chars from the unreserved set. */
export function generateCodeVerifier(length = 128): string {
  if (length < 43 || length > 128) {
    throw new Error("PKCE verifier length must be 43-128");
  }
  const random = new Uint8Array(length);
  crypto.getRandomValues(random);
  let out = "";
  for (const byte of random) {
    out += VERIFIER_CHARS[byte % VERIFIER_CHARS.length];
  }
  return out;
}

/** Base64url without padding, per RFC 7636 appendix A. */
export function base64UrlEncode(bytes: ArrayBuffer): string {
  const bin = String.fromCharCode(...new Uint8Array(bytes));
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

/** S256 challenge = BASE64URL(SHA256(ASCII(verifier))). */
export async function computeCodeChallenge(verifier: string): Promise<string> {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(verifier),
  );
  return base64UrlEncode(digest);
}

/** Random `state` parameter to bind the auth round-trip. */
export function generateState(): string {
  return generateCodeVerifier(43);
}

export interface AuthorizeUrlParams {
  orgUrl: string;
  clientId: string;
  redirectUri: string;
  codeChallenge: string;
  state: string;
  scopes?: string[];
}

export function buildAuthorizeUrl(p: AuthorizeUrlParams): string {
  const scopes = p.scopes ?? ["api", "refresh_token", "openid"];
  const params = new URLSearchParams({
    response_type: "code",
    client_id: p.clientId,
    redirect_uri: p.redirectUri,
    code_challenge: p.codeChallenge,
    code_challenge_method: "S256",
    state: p.state,
    scope: scopes.join(" "),
  });
  return `${p.orgUrl}/services/oauth2/authorize?${params.toString()}`;
}
