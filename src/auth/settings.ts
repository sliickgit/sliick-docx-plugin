/**
 * Add-in settings: which org to talk to, which ECA consumer key to use, and
 * whether to run in mock mode (no Salesforce; demo/dev against canned data).
 *
 * Storage: localStorage. The task pane runs in a persistent webview partition
 * per add-in, so localStorage is appropriate for non-secret settings. Tokens
 * are kept in sessionStorage (see auth.ts) to limit their lifetime on disk.
 */

export interface AddinSettings {
  /** e.g. https://mydomain.my.salesforce.com — empty string until configured */
  orgUrl: string;
  /** ECA consumer key (public client — not a secret) */
  clientId: string;
  /** Run against the in-memory mock API instead of a real org */
  mockMode: boolean;
}

// Bumped to v3 for the per-customer ECA model: each customer creates their own
// local External Client App and enters its consumer key, so any previously
// saved (shared/baked) key must be discarded.
const STORAGE_KEY = "sliick.settings.v3";

/**
 * Shared Sliick global ECA consumer key (public client — not a secret). Used as
 * a fallback when a customer leaves the Settings consumer-key field blank, so
 * the add-in supports both models:
 *   - customer creates their own local ECA and pastes its key, or
 *   - customer leaves it blank and uses the shared Sliick app.
 */
const SLIICK_GLOBAL_CLIENT_ID =
  "3MVG9QJ.PEcCek9ZS2UpB7gXr_1tcAtTAMfQHj0OfVWZ.BChmARUuQ4.cuY1QXbgONr_6IYt1zrcSOwiHlhcx";

// clientId defaults to blank so the field starts empty; effectiveClientId()
// falls back to the shared key when the customer doesn't supply their own.
export const DEFAULT_SETTINGS: AddinSettings = {
  orgUrl: "",
  clientId: "",
  mockMode: true,
};

/** The consumer key to actually use for OAuth: the customer's, or the shared fallback. */
export function effectiveClientId(settings: AddinSettings): string {
  return settings.clientId.trim() || SLIICK_GLOBAL_CLIENT_ID;
}

export function loadSettings(): AddinSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const parsed = JSON.parse(raw) as Partial<AddinSettings>;
    return {
      orgUrl: typeof parsed.orgUrl === "string" ? parsed.orgUrl : "",
      clientId: typeof parsed.clientId === "string" ? parsed.clientId : "",
      mockMode: parsed.mockMode !== false,
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(settings: AddinSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

/** Normalizes an org URL: trims, strips trailing slash, requires https. */
export function normalizeOrgUrl(input: string): string | null {
  const trimmed = input.trim().replace(/\/+$/, "");
  if (!/^https:\/\/[a-z0-9][a-z0-9.-]*\.salesforce\.com$/i.test(trimmed)) {
    return null;
  }
  return trimmed;
}
