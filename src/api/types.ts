/**
 * DTOs for the sliick-docs Office add-in REST contract (v1).
 *
 * Source of truth: sliick-docs/.ai-docs/plan/sprint-office-addin-backend/
 * functional-requirements.md §4. Do not extend these shapes here without
 * updating that contract first — both builds code against it.
 */

// ---------- §4.1 GET /office/v1/objects ----------

export interface SObjectInfo {
  apiName: string;
  label: string;
  custom: boolean;
}

export interface ObjectsResponse {
  objects: SObjectInfo[];
}

// ---------- §4.2 GET /office/v1/discover ----------

export interface MergeFieldDef {
  /** e.g. "Account.Name", "Account.Owner.Email", "RunningUser.Name", "Today" */
  key: string;
  label: string;
  /** Salesforce display type, lowercase: string, currency, date, datetime, boolean, picklist, ... */
  type: string;
}

export interface ChildRelationshipDef {
  relationshipName: string;
  childObjectApiName: string;
  label: string;
}

export interface DiscoverResponse {
  baseObjectApiName: string;
  baseObjectLabel: string;
  rootScalarMergeFields: MergeFieldDef[];
  parentLookupMergeFields: MergeFieldDef[];
  runningUserMergeFields: MergeFieldDef[];
  builtInMergeFields: MergeFieldDef[];
  childRelationships: ChildRelationshipDef[];
}

// ---------- §4.3 GET /office/v1/capabilities ----------

export interface CapabilitiesResponse {
  packageVersion: string;
  grammarVersion: number;
  features: {
    conditionals: boolean;
    inverseConditionals: boolean;
    compoundConditions: boolean;
    childLoops: boolean;
    nestedLoops: boolean;
    aggregates: boolean;
    picklistLabels: boolean;
    imageFields: boolean;
    barcodes: boolean;
    signatureTags: boolean;
    pdfOutput: boolean;
  };
  limits: {
    maxFileMb: number;
    maxParentHops: number;
    maxParentHopsInRepeat: number;
  };
}

// ---------- §4.4 POST /office/v1/templates (step 2 of two-step upload) ----------

export interface SaveTemplateRequest {
  templateId?: string;
  name: string;
  baseObjectApiName: string;
  contentVersionId: string;
  fileName: string;
  testRecordId?: string;
}

export type TagStatus = "Resolved" | "FlsWarning" | "Unresolved" | "Structural";

export interface TagCatalogEntry {
  tag: string;
  status: TagStatus;
  suggestion?: string;
}

export interface SaveTemplateResponse {
  templateId: string;
  versionId: string;
  validationStatus: "Valid" | "Invalid";
  tagCatalog: TagCatalogEntry[];
  warnings: ApiWarning[];
}

export interface ApiWarning {
  code: string;
  message: string;
}

// ---------- §4.5 GET /office/v1/templates ----------

export interface TemplateSummary {
  templateId: string;
  name: string;
  baseObjectApiName: string;
  validationStatus: "Valid" | "Invalid";
  lastModifiedDate: string;
  latestVersionId: string;
  fileName: string;
}

export interface TemplatesListResponse {
  templates: TemplateSummary[];
}

// ---------- §4.6 POST /office/v1/preview ----------

export interface PreviewRequest {
  versionId: string;
  recordId?: string;
}

// Response is a binary .docx (handled as Blob by the client).

// ---------- Errors ----------

export interface ApiError {
  errorCode: string;
  message: string;
  details?: unknown;
}

/** Thrown by clients on non-2xx responses. */
export class SliickApiError extends Error {
  constructor(
    public readonly status: number,
    public readonly body: ApiError,
  ) {
    super(body.message);
    this.name = "SliickApiError";
  }
}
