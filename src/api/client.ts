/**
 * SliickApi — the surface the task pane codes against.
 * Two implementations: RealSliickClient (live org) and MockSliickClient.
 */

import {
  ApiError,
  CapabilitiesResponse,
  DiscoverResponse,
  ObjectsResponse,
  PreviewRequest,
  SaveTemplateRequest,
  SaveTemplateResponse,
  SliickApiError,
  TemplatesListResponse,
} from "./types";

export interface SliickApi {
  getObjects(): Promise<ObjectsResponse>;
  discover(baseObject: string): Promise<DiscoverResponse>;
  getCapabilities(): Promise<CapabilitiesResponse>;
  listTemplates(baseObject?: string): Promise<TemplatesListResponse>;
  /**
   * Two-step upload (§4.4): step 1 creates the ContentVersion via the standard
   * REST API; step 2 attaches + validates via /office/v1/templates.
   * `fileBase64` is the .docx from Office.context.document.getFileAsync.
   */
  saveTemplate(
    req: Omit<SaveTemplateRequest, "contentVersionId">,
    fileBase64: string,
  ): Promise<SaveTemplateResponse>;
  /** Returns the merged .docx as a Blob. */
  preview(req: PreviewRequest): Promise<Blob>;
}

export interface AuthorizedSession {
  /** Org base URL, e.g. https://mydomain.my.salesforce.com */
  instanceUrl: string;
  /** Current bearer token. */
  getAccessToken(): string | null;
  /** Attempt a refresh; returns the new access token or throws. */
  refresh(): Promise<string>;
}

const APEX_BASE = "/services/apexrest/sliick/office/v1";
const DATA_API = "/services/data/v62.0";

export class RealSliickClient implements SliickApi {
  constructor(private readonly session: AuthorizedSession) {}

  getObjects(): Promise<ObjectsResponse> {
    return this.getJson(`${APEX_BASE}/objects`);
  }

  discover(baseObject: string): Promise<DiscoverResponse> {
    return this.getJson(
      `${APEX_BASE}/discover?baseObject=${encodeURIComponent(baseObject)}`,
    );
  }

  getCapabilities(): Promise<CapabilitiesResponse> {
    return this.getJson(`${APEX_BASE}/capabilities`);
  }

  listTemplates(baseObject?: string): Promise<TemplatesListResponse> {
    const qs = baseObject ? `?baseObject=${encodeURIComponent(baseObject)}` : "";
    return this.getJson(`${APEX_BASE}/templates${qs}`);
  }

  async saveTemplate(
    req: Omit<SaveTemplateRequest, "contentVersionId">,
    fileBase64: string,
  ): Promise<SaveTemplateResponse> {
    // Step 1: binary via standard ContentVersion REST (handles large payloads).
    const cv = await this.fetchJson<{ id: string }>(`${DATA_API}/sobjects/ContentVersion`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        Title: req.name,
        PathOnClient: req.fileName,
        VersionData: fileBase64,
      }),
    });
    // Step 2: attach + validate.
    return this.fetchJson<SaveTemplateResponse>(`${APEX_BASE}/templates`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...req, contentVersionId: cv.id }),
    });
  }

  async preview(req: PreviewRequest): Promise<Blob> {
    const resp = await this.authedFetch(`${APEX_BASE}/preview`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(req),
    });
    if (!resp.ok) throw await toApiError(resp);
    return resp.blob();
  }

  // ----- plumbing -----

  private getJson<T>(path: string): Promise<T> {
    return this.fetchJson<T>(path, { method: "GET" });
  }

  private async fetchJson<T>(path: string, init: RequestInit): Promise<T> {
    const resp = await this.authedFetch(path, init);
    if (!resp.ok) throw await toApiError(resp);
    return (await resp.json()) as T;
  }

  /** Bearer-authed fetch with one 401 → refresh → retry. */
  private async authedFetch(path: string, init: RequestInit): Promise<Response> {
    const url = `${this.session.instanceUrl}${path}`;
    const doFetch = (token: string) =>
      fetch(url, {
        ...init,
        headers: { ...(init.headers ?? {}), Authorization: `Bearer ${token}` },
      });

    const token = this.session.getAccessToken();
    if (!token) {
      const fresh = await this.session.refresh();
      return doFetch(fresh);
    }
    const first = await doFetch(token);
    if (first.status !== 401) return first;
    const fresh = await this.session.refresh();
    return doFetch(fresh);
  }
}

async function toApiError(resp: Response): Promise<SliickApiError> {
  let body: ApiError;
  try {
    const json = (await resp.json()) as unknown;
    // Apex REST errors may arrive as our shape or as Salesforce's array shape.
    if (Array.isArray(json) && json.length > 0) {
      const first = json[0] as { errorCode?: string; message?: string };
      body = {
        errorCode: first.errorCode ?? "UNKNOWN",
        message: first.message ?? resp.statusText,
      };
    } else {
      const obj = json as Partial<ApiError>;
      body = {
        errorCode: obj.errorCode ?? "UNKNOWN",
        message: obj.message ?? resp.statusText,
        details: obj.details,
      };
    }
  } catch {
    body = { errorCode: "HTTP_" + resp.status, message: resp.statusText };
  }
  return new SliickApiError(resp.status, body);
}
