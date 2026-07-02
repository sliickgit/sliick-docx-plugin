/**
 * MockSliickClient — in-memory implementation of the v1 contract so the add-in
 * is fully demoable in Word without a Salesforce connection. Behavior mirrors
 * the shipped backend (1.8.0): tag validation statuses, capability flags
 * (nested loops + aggregates on; barcodes + signature tags off), simulated
 * latency.
 */

import { SliickApi } from "./client";
import { parseTag } from "../office/tagParse";
import {
  CapabilitiesResponse,
  DiscoverResponse,
  ObjectsResponse,
  PreviewRequest,
  SaveTemplateRequest,
  SaveTemplateResponse,
  TagCatalogEntry,
  TemplatesListResponse,
  TemplateSummary,
} from "./types";

const LATENCY_MS = 250;

function delay<T>(value: T): Promise<T> {
  return new Promise((resolve) => setTimeout(() => resolve(value), LATENCY_MS));
}

const OBJECTS: ObjectsResponse = {
  objects: [
    { apiName: "Account", label: "Account", custom: false },
    { apiName: "Contact", label: "Contact", custom: false },
    { apiName: "Opportunity", label: "Opportunity", custom: false },
    { apiName: "Case", label: "Case", custom: false },
    { apiName: "Invoice__c", label: "Invoice", custom: true },
  ],
};

const FIELDS: Record<string, DiscoverResponse> = {
  Account: {
    baseObjectApiName: "Account",
    baseObjectLabel: "Account",
    rootScalarMergeFields: [
      { key: "Account.Name", label: "Account Name", type: "string" },
      { key: "Account.AccountNumber", label: "Account Number", type: "string" },
      { key: "Account.Industry", label: "Industry", type: "picklist" },
      { key: "Account.AnnualRevenue", label: "Annual Revenue", type: "currency" },
      { key: "Account.NumberOfEmployees", label: "Employees", type: "int" },
      { key: "Account.Website", label: "Website", type: "url" },
      { key: "Account.Phone", label: "Account Phone", type: "phone" },
      { key: "Account.BillingCity", label: "Billing City", type: "string" },
      { key: "Account.BillingCountry", label: "Billing Country", type: "string" },
      { key: "Account.CreatedDate", label: "Created Date", type: "datetime" },
    ],
    parentLookupMergeFields: [
      { key: "Account.Owner.Name", label: "Owner › Full Name", type: "string" },
      { key: "Account.Owner.Email", label: "Owner › Email", type: "email" },
      { key: "Account.Parent.Name", label: "Parent Account › Name", type: "string" },
    ],
    runningUserMergeFields: [
      { key: "RunningUser.Name", label: "Running User › Name", type: "string" },
      { key: "RunningUser.Email", label: "Running User › Email", type: "email" },
      { key: "RunningUser.Title", label: "Running User › Title", type: "string" },
      { key: "RunningUser.CompanyName", label: "Running User › Company", type: "string" },
    ],
    builtInMergeFields: [
      { key: "Today", label: "Today's Date", type: "date" },
      { key: "Now", label: "Current Date/Time", type: "datetime" },
    ],
    childRelationships: [
      { relationshipName: "Contacts", childObjectApiName: "Contact", label: "Contacts" },
      { relationshipName: "Opportunities", childObjectApiName: "Opportunity", label: "Opportunities" },
      { relationshipName: "Cases", childObjectApiName: "Case", label: "Cases" },
      { relationshipName: "Approvals", childObjectApiName: "ProcessInstanceStep", label: "Approval History" },
    ],
  },
  Contact: {
    baseObjectApiName: "Contact",
    baseObjectLabel: "Contact",
    rootScalarMergeFields: [
      { key: "Contact.FirstName", label: "First Name", type: "string" },
      { key: "Contact.LastName", label: "Last Name", type: "string" },
      { key: "Contact.Email", label: "Email", type: "email" },
      { key: "Contact.Phone", label: "Phone", type: "phone" },
      { key: "Contact.Title", label: "Title", type: "string" },
    ],
    parentLookupMergeFields: [
      { key: "Contact.Account.Name", label: "Account › Name", type: "string" },
      { key: "Contact.Owner.Name", label: "Owner › Name", type: "string" },
    ],
    runningUserMergeFields: [
      { key: "RunningUser.Name", label: "Running User › Name", type: "string" },
      { key: "RunningUser.Email", label: "Running User › Email", type: "email" },
    ],
    builtInMergeFields: [
      { key: "Today", label: "Today's Date", type: "date" },
      { key: "Now", label: "Current Date/Time", type: "datetime" },
    ],
    childRelationships: [
      { relationshipName: "Cases", childObjectApiName: "Case", label: "Cases" },
      { relationshipName: "Approvals", childObjectApiName: "ProcessInstanceStep", label: "Approval History" },
    ],
  },
  Opportunity: {
    baseObjectApiName: "Opportunity",
    baseObjectLabel: "Opportunity",
    rootScalarMergeFields: [
      { key: "Opportunity.Name", label: "Opportunity Name", type: "string" },
      { key: "Opportunity.Amount", label: "Amount", type: "currency" },
      { key: "Opportunity.StageName", label: "Stage", type: "picklist" },
      { key: "Opportunity.CloseDate", label: "Close Date", type: "date" },
      { key: "Opportunity.Probability", label: "Probability", type: "percent" },
    ],
    parentLookupMergeFields: [
      { key: "Opportunity.Account.Name", label: "Account › Name", type: "string" },
      { key: "Opportunity.Owner.Name", label: "Owner › Name", type: "string" },
    ],
    runningUserMergeFields: [
      { key: "RunningUser.Name", label: "Running User › Name", type: "string" },
    ],
    builtInMergeFields: [
      { key: "Today", label: "Today's Date", type: "date" },
    ],
    childRelationships: [
      {
        relationshipName: "OpportunityLineItems",
        childObjectApiName: "OpportunityLineItem",
        label: "Products (Line Items)",
      },
      { relationshipName: "Approvals", childObjectApiName: "ProcessInstanceStep", label: "Approval History" },
    ],
  },
  Case: {
    baseObjectApiName: "Case",
    baseObjectLabel: "Case",
    rootScalarMergeFields: [
      { key: "Case.CaseNumber", label: "Case Number", type: "string" },
      { key: "Case.Subject", label: "Subject", type: "string" },
      { key: "Case.Status", label: "Status", type: "picklist" },
      { key: "Case.Priority", label: "Priority", type: "picklist" },
      { key: "Case.CreatedDate", label: "Created Date", type: "datetime" },
    ],
    parentLookupMergeFields: [
      { key: "Case.Account.Name", label: "Account › Name", type: "string" },
      { key: "Case.Contact.Name", label: "Contact › Name", type: "string" },
    ],
    runningUserMergeFields: [
      { key: "RunningUser.Name", label: "Running User › Name", type: "string" },
    ],
    builtInMergeFields: [
      { key: "Today", label: "Today's Date", type: "date" },
    ],
    childRelationships: [],
  },
  OpportunityLineItem: {
    baseObjectApiName: "OpportunityLineItem",
    baseObjectLabel: "Opportunity Product",
    rootScalarMergeFields: [
      { key: "OpportunityLineItem.Name", label: "Product Name", type: "string" },
      { key: "OpportunityLineItem.Quantity", label: "Quantity", type: "double" },
      { key: "OpportunityLineItem.UnitPrice", label: "Unit Price", type: "currency" },
      { key: "OpportunityLineItem.TotalPrice", label: "Total Price", type: "currency" },
    ],
    parentLookupMergeFields: [],
    runningUserMergeFields: [],
    builtInMergeFields: [],
    childRelationships: [],
  },
};

/**
 * In-loop field keys for the synthetic Approvals relationship — mirrors
 * TemplateMergeFieldService.APPROVALS_FIELDS in sliick-docs.
 */
const APPROVALS_KEYS = new Set([
  "ActorName",
  "ActorTitle",
  "StepStatus",
  "Comments",
  "ActedAt",
  "ProcessName",
  "StepName",
]);

const CAPABILITIES: CapabilitiesResponse = {
  packageVersion: "mock-1.8.0",
  grammarVersion: 1,
  features: {
    conditionals: true,
    inverseConditionals: true,
    compoundConditions: true,
    childLoops: true,
    nestedLoops: true,
    aggregates: true,
    picklistLabels: true,
    imageFields: true,
    barcodes: true,
    signatureTags: false,
    pdfOutput: true,
    // grammar-v2 (mirrors OfficeAddinService.getCapabilities post-5dca571)
    loopFilters: true,
    fallbackText: true,
    localeFormats: true,
  },
  limits: { maxFileMb: 10, maxParentHops: 5, maxParentHopsInRepeat: 1 },
};

/** Tags resolvable in mock mode = everything discover() returns for the object. */
function knownKeys(d: DiscoverResponse): Set<string> {
  const keys = new Set<string>();
  for (const list of [
    d.rootScalarMergeFields,
    d.parentLookupMergeFields,
    d.runningUserMergeFields,
    d.builtInMergeFields,
  ]) {
    for (const f of list) keys.add(f.key);
  }
  return keys;
}

/**
 * Classifies the inner text of one {{...}} tag the way the backend validator
 * documents it. Delegates to the shared grammar parser (tagParse), so the
 * mock understands grammar-v2 forms (loop WHERE/ORDER BY, fallback pipes,
 * barcodes) exactly like the live engine. Exported for unit tests.
 */
export function classifyTag(inner: string, resolvable: Set<string>): TagCatalogEntry {
  const tag = `{{${inner}}}`;
  const parsed = parseTag(inner);
  switch (parsed.kind) {
    case "ifOpen":
    case "elseMarker":
    case "ifClose":
    case "inverseOpen":
    case "loopOpen":
    case "blockClose":
    case "signature":
      return { tag, status: "Structural" };
    case "aggregate":
      // Mock accepts any well-formed aggregate; the backend validates the rel/field.
      return { tag, status: "Resolved" };
    case "scalar":
    case "image":
    case "barcode": {
      if (resolvable.has(parsed.key)) {
        return { tag, status: "Resolved" };
      }
      const suggestion = closestKey(parsed.key, resolvable);
      return suggestion
        ? { tag, status: "Unresolved", suggestion }
        : { tag, status: "Unresolved" };
    }
    default:
      return { tag, status: "Unresolved" };
  }
}

/** Cheap fuzzy suggestion: case-insensitive match or shared-prefix candidate. */
function closestKey(input: string, keys: Set<string>): string | undefined {
  const lower = input.toLowerCase();
  for (const k of keys) {
    if (k.toLowerCase() === lower) return k;
  }
  let best: string | undefined;
  let bestLen = 3; // require at least a 4-char shared prefix
  for (const k of keys) {
    const shared = sharedPrefixLength(k.toLowerCase(), lower);
    if (shared > bestLen) {
      bestLen = shared;
      best = k;
    }
  }
  return best;
}

function sharedPrefixLength(a: string, b: string): number {
  const max = Math.min(a.length, b.length);
  let i = 0;
  while (i < max && a[i] === b[i]) i++;
  return i;
}

export function extractTags(text: string): string[] {
  const out: string[] = [];
  const re = /\{\{([^{}]+)\}\}/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    out.push(m[1] as string);
  }
  return out;
}

/** In-loop resolvable keys for a child relationship ("FirstName", not "Contact.FirstName"). */
function inLoopKeys(baseObject: string, relationshipName: string): Set<string> | null {
  if (relationshipName === "Approvals") return APPROVALS_KEYS;
  const base = FIELDS[baseObject];
  const rel = base?.childRelationships.find(
    (r) => r.relationshipName === relationshipName,
  );
  if (!rel) return null; // unknown relationship — be lenient in mock mode
  const child = FIELDS[rel.childObjectApiName];
  if (!child) return null;
  const keys = new Set<string>();
  for (const f of child.rootScalarMergeFields) {
    const dot = f.key.indexOf(".");
    keys.add(dot === -1 ? f.key : f.key.slice(dot + 1));
  }
  return keys;
}

/**
 * Scope-aware classification of a whole document's tags, mirroring the backend
 * validator: inside `{{#Rel}}…{{/Rel}}` bare field keys resolve against the
 * child object's fields. Exported for unit tests.
 */
export function classifyDocumentTags(
  text: string,
  baseObject: string,
): TagCatalogEntry[] {
  const discover = FIELDS[baseObject];
  const rootResolvable = discover ? knownKeys(discover) : new Set<string>();
  const loopStack: string[] = [];
  const out: TagCatalogEntry[] = [];

  for (const inner of extractTags(text)) {
    const tag = `{{${inner}}}`;
    const parsed = parseTag(inner);

    if (
      parsed.kind === "ifOpen" ||
      parsed.kind === "elseMarker" ||
      parsed.kind === "ifClose" ||
      parsed.kind === "inverseOpen" ||
      parsed.kind === "signature"
    ) {
      out.push({ tag, status: "Structural" });
      continue;
    }
    if (parsed.kind === "loopOpen") {
      // Relationship name only — WHERE/ORDER BY modifiers don't change scope.
      loopStack.push(parsed.relationship);
      out.push({ tag, status: "Structural" });
      continue;
    }
    if (parsed.kind === "blockClose") {
      // Pop only a matching loop close — {{/Field}} closing an inverse block
      // must not eject the enclosing loop scope.
      if (loopStack[loopStack.length - 1] === parsed.target) {
        loopStack.pop();
      }
      out.push({ tag, status: "Structural" });
      continue;
    }

    const currentLoop = loopStack[loopStack.length - 1];
    if (currentLoop !== undefined) {
      const keys = inLoopKeys(baseObject, currentLoop);
      if (keys === null) {
        out.push({ tag, status: "Resolved" }); // unknown child schema — lenient
      } else {
        out.push(classifyTag(inner, keys));
      }
      continue;
    }
    out.push(classifyTag(inner, rootResolvable));
  }
  return out;
}

export class MockSliickClient implements SliickApi {
  private readonly saved: TemplateSummary[] = [];
  private counter = 0;
  /** Set by the task pane before save so the mock can lint real document text. */
  documentText = "";

  getObjects(): Promise<ObjectsResponse> {
    return delay(OBJECTS);
  }

  discover(baseObject: string): Promise<DiscoverResponse> {
    const found = FIELDS[baseObject];
    if (!found) {
      return Promise.reject(new Error(`Mock has no field data for ${baseObject}`));
    }
    return delay(found);
  }

  getCapabilities(): Promise<CapabilitiesResponse> {
    return delay(CAPABILITIES);
  }

  listTemplates(baseObject?: string): Promise<TemplatesListResponse> {
    const templates = baseObject
      ? this.saved.filter((t) => t.baseObjectApiName === baseObject)
      : [...this.saved];
    return delay({ templates });
  }

  saveTemplate(
    req: Omit<SaveTemplateRequest, "contentVersionId">,
    _fileBase64: string,
  ): Promise<SaveTemplateResponse> {
    this.counter += 1;
    const tagCatalog = classifyDocumentTags(this.documentText, req.baseObjectApiName);
    const hasUnresolved = tagCatalog.some((t) => t.status === "Unresolved");
    // Revise-by-id (new version of an existing template) keeps the templateId;
    // otherwise fall back to upsert-by-name, matching the backend's behavior.
    const existing = req.templateId
      ? this.saved.findIndex((t) => t.templateId === req.templateId)
      : this.saved.findIndex((t) => t.name === req.name);
    if (req.templateId && existing < 0) {
      return Promise.reject(new Error(`Mock has no template ${req.templateId}`));
    }
    const summary: TemplateSummary = {
      templateId: existing >= 0 ? this.saved[existing]!.templateId : `mockT${this.counter}`,
      name: req.name,
      baseObjectApiName: req.baseObjectApiName,
      validationStatus: hasUnresolved ? "Invalid" : "Valid",
      lastModifiedDate: new Date(Date.now()).toISOString(),
      latestVersionId: `mockV${this.counter}`,
      fileName: req.fileName,
    };
    if (existing >= 0) this.saved[existing] = summary;
    else this.saved.push(summary);

    return delay({
      templateId: summary.templateId,
      versionId: summary.latestVersionId,
      validationStatus: summary.validationStatus,
      tagCatalog,
      warnings:
        tagCatalog.length === 0
          ? [{ code: "NO_TAGS", message: "No merge tags found in the document." }]
          : [],
      // The real native-PDF verdict comes from the backend OOXML lint (text
      // boxes, fonts, …), which can't be detected from plain body text — mock
      // mode reports PDF-ready so the badge is demoable.
      pdfReady: true,
      pdfWarnings: [],
    });
  }

  preview(_req: PreviewRequest): Promise<Blob> {
    return Promise.reject(
      new Error("Preview requires a Salesforce connection (not available in mock mode)."),
    );
  }
}
