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

const STORAGE_KEY = "sliick.settings.v1";

export const DEFAULT_SETTINGS: AddinSettings = {
  orgUrl: "",
  clientId: "",
  mockMode: true,
};

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
