/**
 * Sliick Docs task pane — application shell.
 *
 * Single-state-object + full re-render. The pane is small enough that a
 * framework would cost more than it buys; all DOM writes go through render().
 */

import { RealSliickClient, SliickApi, AuthorizedSession } from "../api/client";
import { MockSliickClient } from "../api/mock";
import {
  CapabilitiesResponse,
  ChildRelationshipDef,
  DiscoverResponse,
  MergeFieldDef,
  SaveTemplateResponse,
  SObjectInfo,
  TemplateSummary,
} from "../api/types";
import {
  clearTokens,
  getStoredTokens,
  login,
  refreshAccessToken,
  TokenSet,
} from "../auth/auth";
import {
  AddinSettings,
  effectiveClientId,
  loadSettings,
  normalizeOrgUrl,
  saveSettings,
} from "../auth/settings";
import {
  AggregateFn,
  aggregateTag,
  barcodeTag,
  compoundConditionTags,
  ConditionClause,
  ConditionOperator,
  conditionExpressionText,
  defaultFormatForType,
  imageTag,
  inLoopFieldKey,
  LOCALE_KEYWORD_FORMATS,
  LoopModifiers,
  loopRowCellTexts,
  loopTags,
  scalarTag,
  scalarTagWithOptions,
} from "../office/tags";
import {
  APPROVALS_LOOP_FIELDS,
  buildLintContext,
  LintOutcome,
  lintDocumentText,
} from "../office/lint";
import { enclosingLoopAt, findTags, ParsedTag } from "../office/tagParse";
import {
  getDocumentAsBase64,
  getDocumentText,
  getSelectionContext,
  insertConditional,
  insertConditionalTags,
  insertInverse,
  insertLoopTable,
  insertNestedLoopWithTable,
  insertParagraphBlock,
  insertScalar,
  selectTagOccurrence,
  setTagHighlights,
  suggestedTemplateName,
} from "../office/wordInsert";

// ---------------------------------------------------------------- state

type View =
  | { kind: "main" }
  | { kind: "settings" }
  | { kind: "connect" }
  | {
      kind: "loopWizard";
      rel: ChildRelationshipDef;
      childFields: MergeFieldDef[];
      /** The child object's own child relationships — nested-loop candidates. */
      childRels: ChildRelationshipDef[];
      /** Checked column indices — kept in state so re-renders don't reset them. */
      childChecked: number[];
      nestedRel?: ChildRelationshipDef;
      nestedFields?: MergeFieldDef[];
      nestedChecked?: number[];
    }
  | { kind: "condWizard"; mode: "if" | "inverse" }
  | { kind: "imageWizard" }
  | { kind: "barcodeWizard"; initial?: { key: string; barcodeType: "code128" | "qr"; size?: string } }
  | { kind: "fieldOptions"; field: MergeFieldDef; initial?: { format?: string; locale?: string; fallback?: string } }
  | { kind: "editLoop"; relationship: string; where?: string; orderBy?: string; descending?: boolean }
  | { kind: "templates"; templates: TemplateSummary[] }
  | { kind: "tags" }
  | { kind: "save" }
  | { kind: "lintResults"; result: SaveTemplateResponse };

interface AppState {
  settings: AddinSettings;
  api: SliickApi | null;
  mock: MockSliickClient | null;
  tokens: TokenSet | null;
  capabilities: CapabilitiesResponse | null;
  objects: SObjectInfo[];
  baseObject: string | null;
  discover: DiscoverResponse | null;
  search: string;
  collapsed: Set<string>;
  view: View;
  busy: string | null;
  error: string | null;
  notice: string | null;
  lastSavedVersionId: string | null;
  /** Set when the next save should create a new version of an existing template. */
  activeTemplate: { templateId: string; name: string } | null;
  /** Test record Id from the last save — prefills the ad-hoc preview input. */
  lastTestRecordId: string;
  /** Latest local lint run (tags panel). */
  lint: LintOutcome | null;
  /** Child discovers fetched for lint/scope, keyed by childObjectApiName. */
  childDiscovers: Map<string, DiscoverResponse>;
  /** Authoring-only tag highlighting toggle (stripped automatically on save). */
  highlightsOn: boolean;
  /** Tag under the cursor/selection that a wizard can edit. */
  editTag: { tag: string; occurrence: number; parsed: ParsedTag } | null;
  /** Relationship of the loop the cursor sits inside, or null at root scope. */
  cursorLoop: string | null;
  /** When set, the next wizard insert replaces this tag instead of inserting fresh. */
  replaceTarget: { tag: string; occurrence: number } | null;
}

const state: AppState = {
  settings: loadSettings(),
  api: null,
  mock: null,
  tokens: null,
  capabilities: null,
  objects: [],
  baseObject: null,
  discover: null,
  search: "",
  collapsed: new Set(),
  view: { kind: "main" },
  busy: null,
  error: null,
  notice: null,
  lastSavedVersionId: null,
  activeTemplate: null,
  lastTestRecordId: "",
  lint: null,
  childDiscovers: new Map(),
  highlightsOn: false,
  editTag: null,
  cursorLoop: null,
  replaceTarget: null,
};

// ---------------------------------------------------------------- boot

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    document.getElementById("app")!.innerHTML =
      `<div class="pane"><div class="banner error">Sliick Docs currently supports Word. Excel and PowerPoint are coming next.</div></div>`;
    return;
  }
  // Selection smarts: debounce-probe the cursor for editable tags + loop scope.
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    scheduleSelectionProbe,
  );
  void initApi();
});

let selectionTimer: number | undefined;
function scheduleSelectionProbe(): void {
  if (selectionTimer !== undefined) window.clearTimeout(selectionTimer);
  selectionTimer = window.setTimeout(() => void probeSelection(), 350);
}

const EDITABLE_TAG_KINDS = new Set(["scalar", "barcode", "loopOpen"]);

/**
 * Size cutoff for the selection probe (~150+ pages of text). Each probe reads
 * the whole document text, so on very large documents the smart-cursor
 * features are quietly disabled after the first oversized read — graceful
 * absence beats an editor that hiccups on every pause. Session-sticky:
 * documents don't meaningfully shrink mid-session.
 */
const MAX_PROBE_TEXT_CHARS = 400_000;
let probeDisabledForSize = false;

/**
 * Reads the selection, hit-tests it against the document's tags, and detects
 * the enclosing loop scope. Renders only when something changed, and only on
 * the main view — wizards must not be re-rendered under the user's fingers.
 */
async function probeSelection(): Promise<void> {
  if (probeDisabledForSize) return;
  if (state.view.kind !== "main" || !state.discover || state.busy) return;
  try {
    const { docText, cursorStart, cursorEnd } = await getSelectionContext();
    if (docText.length > MAX_PROBE_TEXT_CHARS) {
      probeDisabledForSize = true;
      if (state.editTag || state.cursorLoop) {
        state.editTag = null;
        state.cursorLoop = null;
        if (state.view.kind === "main") render();
      }
      return;
    }
    const tags = findTags(docText);
    const hit =
      tags.find((t) => t.index < cursorEnd && cursorStart < t.index + t.tag.length) ??
      tags.find((t) => t.index <= cursorStart && cursorStart <= t.index + t.tag.length);
    const editable = hit && EDITABLE_TAG_KINDS.has(hit.parsed.kind) ? hit : undefined;
    const loopRel = enclosingLoopAt(docText, cursorStart);

    const tagChanged =
      (state.editTag?.tag ?? null) !== (editable?.tag ?? null) ||
      (state.editTag?.occurrence ?? -1) !== (editable?.occurrence ?? -1);
    const loopChanged = state.cursorLoop !== loopRel;
    if (!tagChanged && !loopChanged) return;

    state.editTag = editable
      ? { tag: editable.tag, occurrence: editable.occurrence, parsed: editable.parsed }
      : null;
    state.cursorLoop = loopRel;
    if (loopRel) await ensureScopeFieldsFor(loopRel);
    if (state.view.kind === "main") render();
  } catch {
    // Selection probing is best-effort; never surface errors for it.
  }
}

/** Fetches (caches) the child discover backing a loop-scope field list. */
async function ensureScopeFieldsFor(relationshipName: string): Promise<void> {
  const d = state.discover;
  if (!d || !state.api) return;
  const rel = d.childRelationships.find(
    (r) => r.relationshipName.toLowerCase() === relationshipName.toLowerCase(),
  );
  if (!rel || rel.relationshipName === "Approvals") return;
  if (state.childDiscovers.has(rel.childObjectApiName)) return;
  try {
    state.childDiscovers.set(rel.childObjectApiName, await state.api.discover(rel.childObjectApiName));
  } catch {
    // Unknown child schema — the scope section simply won't show.
  }
}

/** In-loop fields for a relationship the cursor sits inside, if known. */
function scopeFieldsFor(relationshipName: string): { label: string; fields: MergeFieldDef[] } | null {
  const d = state.discover;
  if (!d) return null;
  const rel = d.childRelationships.find(
    (r) => r.relationshipName.toLowerCase() === relationshipName.toLowerCase(),
  );
  if (!rel) return null;
  if (rel.relationshipName === "Approvals") {
    return { label: rel.label, fields: APPROVALS_LOOP_FIELDS };
  }
  const child = state.childDiscovers.get(rel.childObjectApiName);
  if (!child) return null;
  return {
    label: rel.label,
    fields: child.rootScalarMergeFields.map((f) => ({ ...f, key: inLoopFieldKey(f.key) })),
  };
}

/** Replaces the edit-target tag when set (select → insert), else inserts fresh. */
async function insertOrReplaceScalar(tagText: string): Promise<void> {
  if (state.replaceTarget) {
    await selectTagOccurrence(state.replaceTarget.tag, state.replaceTarget.occurrence);
    state.replaceTarget = null;
  }
  await insertScalar(tagText);
}

/** Opens the wizard matching the selected tag, prefilled, in replace mode. */
function openTagEditor(target: { tag: string; occurrence: number; parsed: ParsedTag }): void {
  const p = target.parsed;
  state.replaceTarget = { tag: target.tag, occurrence: target.occurrence };
  if (p.kind === "scalar") {
    const known = [
      ...(state.discover?.rootScalarMergeFields ?? []),
      ...(state.discover?.parentLookupMergeFields ?? []),
      ...(state.discover?.runningUserMergeFields ?? []),
      ...(state.discover?.builtInMergeFields ?? []),
    ].find((f) => f.key === p.key);
    state.view = {
      kind: "fieldOptions",
      field: known ?? { key: p.key, label: p.key },
      initial: { format: p.format, locale: p.locale, fallback: p.fallback },
    };
  } else if (p.kind === "barcode") {
    state.view = {
      kind: "barcodeWizard",
      initial: { key: p.key, barcodeType: p.barcodeType, size: p.size },
    };
  } else if (p.kind === "loopOpen") {
    state.view = {
      kind: "editLoop",
      relationship: p.relationship,
      where: p.where,
      orderBy: p.orderBy,
      descending: p.descending,
    };
  } else {
    state.replaceTarget = null;
    return;
  }
  render();
}

async function initApi(): Promise<void> {
  state.error = null;
  if (state.settings.mockMode) {
    state.mock = new MockSliickClient();
    state.api = state.mock;
    state.view = { kind: "main" };
  } else {
    state.mock = null;
    state.tokens = getStoredTokens();
    if (!state.tokens) {
      state.api = null;
      state.view = { kind: "connect" };
      render();
      return;
    }
    state.api = new RealSliickClient(makeSession());
    state.view = { kind: "main" };
  }
  await loadCatalog();
}

function makeSession(): AuthorizedSession {
  return {
    instanceUrl: state.tokens?.instanceUrl ?? state.settings.orgUrl,
    getAccessToken: () => state.tokens?.accessToken ?? null,
    refresh: async () => {
      state.tokens = await refreshAccessToken(
        state.settings.orgUrl,
        effectiveClientId(state.settings),
      );
      return state.tokens.accessToken;
    },
  };
}

async function loadCatalog(): Promise<void> {
  if (!state.api) return;
  await withBusy("Loading Salesforce metadata…", async () => {
    const [objects, capabilities] = await Promise.all([
      state.api!.getObjects(),
      state.api!.getCapabilities(),
    ]);
    state.objects = objects.objects;
    state.capabilities = capabilities;
    if (!state.baseObject && state.objects.length > 0) {
      state.baseObject = state.objects[0]!.apiName;
    }
    if (state.baseObject) {
      state.discover = await state.api!.discover(state.baseObject);
    }
  });
}

async function selectBaseObject(apiName: string): Promise<void> {
  state.baseObject = apiName;
  state.discover = null;
  state.activeTemplate = null; // templates are per-base-object
  await withBusy(`Loading ${apiName} fields…`, async () => {
    state.discover = await state.api!.discover(apiName);
  });
}

// ---------------------------------------------------------------- actions

async function withBusy(message: string | null, fn: () => Promise<void>): Promise<void> {
  state.busy = message;
  state.error = null;
  render();
  try {
    await fn();
  } catch (e) {
    state.error = friendlyError(e);
  } finally {
    state.busy = null;
    render();
  }
}

/** Maps raw auth/network failures to actionable guidance. */
function friendlyError(e: unknown): string {
  const message = e instanceof Error ? e.message : String(e);
  const status = (e as { status?: number }).status;
  if (status === 401 || /unauthorized|invalid_grant|session expired|INVALID_SESSION/i.test(message)) {
    return "Your Salesforce session has expired — sign in again (Settings ⚙ → Sign out, then Connect).";
  }
  if (/failed to fetch|networkerror|load failed/i.test(message)) {
    return "Can't reach Salesforce — check your connection, the org URL in Settings, and the org's CORS setup.";
  }
  return message;
}

// ---------- recents + pins (per base object, localStorage) ----------

const RECENTS_LIMIT = 6;
const recentsKey = (): string => `sliick.recent.${state.baseObject ?? ""}`;
const pinsKey = (): string => `sliick.pins.${state.baseObject ?? ""}`;

function loadKeyList(storageKey: string): string[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(storageKey) ?? "[]") as unknown;
    return Array.isArray(parsed) ? parsed.filter((x): x is string => typeof x === "string") : [];
  } catch {
    return [];
  }
}

function saveKeyList(storageKey: string, keys: string[]): void {
  try {
    localStorage.setItem(storageKey, JSON.stringify(keys));
  } catch {
    // storage unavailable — recents/pins silently off
  }
}

function noteRecentField(fieldKey: string): void {
  const keys = loadKeyList(recentsKey()).filter((k) => k !== fieldKey);
  keys.unshift(fieldKey);
  saveKeyList(recentsKey(), keys.slice(0, RECENTS_LIMIT));
}

function isPinned(fieldKey: string): boolean {
  return loadKeyList(pinsKey()).includes(fieldKey);
}

function togglePin(fieldKey: string): void {
  const keys = loadKeyList(pinsKey());
  const at = keys.indexOf(fieldKey);
  if (at === -1) keys.push(fieldKey);
  else keys.splice(at, 1);
  saveKeyList(pinsKey(), keys);
}

async function onInsertField(field: MergeFieldDef): Promise<void> {
  const format = defaultFormatForType(field.type);
  noteRecentField(field.key);
  await withBusy(null, async () => {
    await insertScalar(scalarTag(field.key, format));
  });
}

/** Default column pick: the first three fields (or fewer). */
function firstThreeIndices(fields: MergeFieldDef[]): number[] {
  return fields.slice(0, 3).map((_, i) => i);
}

async function onLoopWizardOpen(rel: ChildRelationshipDef): Promise<void> {
  if (rel.relationshipName === "Approvals") {
    state.view = {
      kind: "loopWizard",
      rel,
      childFields: APPROVALS_LOOP_FIELDS,
      childRels: [],
      childChecked: firstThreeIndices(APPROVALS_LOOP_FIELDS),
    };
    render();
    return;
  }
  await withBusy(`Loading ${rel.label} fields…`, async () => {
    const childDiscover = await state.api!.discover(rel.childObjectApiName);
    state.view = {
      kind: "loopWizard",
      rel,
      childFields: childDiscover.rootScalarMergeFields,
      // Nested-loop candidates. Approvals is synthetic and only resolves
      // against the base record, so it can't be a nested (grandchild) source.
      childRels: childDiscover.childRelationships.filter(
        (r) => r.relationshipName !== "Approvals",
      ),
      childChecked: firstThreeIndices(childDiscover.rootScalarMergeFields),
    };
  });
}

async function onNestedRelPick(rel: ChildRelationshipDef | undefined): Promise<void> {
  if (state.view.kind !== "loopWizard") return;
  const view = state.view;
  if (!rel) {
    state.view = { ...view, nestedRel: undefined, nestedFields: undefined, nestedChecked: undefined };
    render();
    return;
  }
  await withBusy(`Loading ${rel.label} fields…`, async () => {
    const grandDiscover = await state.api!.discover(rel.childObjectApiName);
    state.view = {
      ...view,
      nestedRel: rel,
      nestedFields: grandDiscover.rootScalarMergeFields,
      nestedChecked: firstThreeIndices(grandDiscover.rootScalarMergeFields),
    };
  });
}

async function onNestedLoopInsert(
  rel: ChildRelationshipDef,
  childPicked: MergeFieldDef[],
  nestedRel: ChildRelationshipDef,
  nestedPicked: MergeFieldDef[],
  modifiers?: LoopModifiers,
): Promise<void> {
  // Titan shape: outer loop paragraphs wrapping a REAL table whose data row
  // is a row-scope inner loop (engine-covered invoice layout).
  const parentLine = childPicked
    .map((f) => scalarTag(inLoopFieldKey(f.key), defaultFormatForType(f.type)))
    .join(" — ");
  const headers = nestedPicked.map((f) => f.label);
  const cellKeys = nestedPicked.map((f) => {
    const key = inLoopFieldKey(f.key);
    const format = defaultFormatForType(f.type);
    return format ? `${key}:${format}` : key;
  });
  const outer = loopTags(rel.relationshipName, modifiers);
  const dataRow = loopRowCellTexts(nestedRel.relationshipName, cellKeys);
  await withBusy("Inserting nested table…", async () => {
    await insertNestedLoopWithTable(outer.open, parentLine, headers, dataRow, outer.close);
    state.view = { kind: "main" };
  });
}

async function onLoopInsert(
  rel: ChildRelationshipDef,
  picked: MergeFieldDef[],
  modifiers?: LoopModifiers,
): Promise<void> {
  await withBusy("Inserting table…", async () => {
    await insertLoopTable(
      rel.relationshipName,
      picked.map((f) => ({ inLoopKey: inLoopFieldKey(f.key), label: f.label })),
      modifiers,
    );
    state.view = { kind: "main" };
  });
}

async function onSave(name: string, testRecordId: string): Promise<void> {
  await withBusy("Saving to Salesforce…", async () => {
    // Authoring highlights must never reach the saved template — strip every
    // tag highlight before the document bytes are captured.
    try {
      await setTagHighlights(await documentTagTexts(), null);
      state.highlightsOn = false;
    } catch {
      // Non-fatal: worst case the highlight survives; the doc still saves.
    }
    if (state.mock) {
      state.mock.documentText = await getDocumentText();
    }
    const fileBase64 = await getDocumentAsBase64();
    const sizeMb = (fileBase64.length * 3) / 4 / (1024 * 1024);
    const maxMb = state.capabilities?.limits.maxFileMb ?? 10;
    if (sizeMb > maxMb) {
      throw new Error(
        `Document is ${sizeMb.toFixed(1)} MB — over the ${maxMb} MB template limit. Compress images (Picture Format → Compress Pictures) and try again.`,
      );
    }
    const result = await state.api!.saveTemplate(
      {
        name,
        baseObjectApiName: state.baseObject!,
        fileName: `${name}.docx`,
        ...(state.activeTemplate ? { templateId: state.activeTemplate.templateId } : {}),
        ...(testRecordId.trim() ? { testRecordId: testRecordId.trim() } : {}),
      },
      fileBase64,
    );
    state.lastSavedVersionId = result.versionId;
    state.lastTestRecordId = testRecordId.trim();
    // Follow-up saves from this pane revise the same template (new versions)
    // instead of accumulating name-collision duplicates.
    state.activeTemplate = { templateId: result.templateId, name };
    state.view = { kind: "lintResults", result };
  });
}

async function onTemplatesOpen(): Promise<void> {
  await withBusy("Loading templates…", async () => {
    const resp = await state.api!.listTemplates(state.baseObject ?? undefined);
    state.view = { kind: "templates", templates: resp.templates };
  });
}

/** Fetches (and caches) child discovers for every loop the document references. */
async function ensureChildDiscoversFor(text: string): Promise<void> {
  const d = state.discover;
  if (!d || !state.api) return;
  const wanted = new Set<string>();
  for (const found of findTags(text)) {
    if (found.parsed.kind === "loopOpen") wanted.add(found.parsed.relationship.toLowerCase());
    if (found.parsed.kind === "aggregate") wanted.add(found.parsed.relationship.toLowerCase());
  }
  for (const rel of d.childRelationships) {
    if (!wanted.has(rel.relationshipName.toLowerCase())) continue;
    if (rel.relationshipName === "Approvals") continue; // fixed field set, no fetch
    if (state.childDiscovers.has(rel.childObjectApiName)) continue;
    try {
      state.childDiscovers.set(rel.childObjectApiName, await state.api.discover(rel.childObjectApiName));
    } catch {
      // Unknown child schema — lint stays lenient for this relationship.
    }
  }
}

/** Runs the local tag check and opens the tags panel. */
async function onReviewTags(): Promise<void> {
  if (!state.discover) return;
  await withBusy("Checking tags…", async () => {
    const text = await getDocumentText();
    await ensureChildDiscoversFor(text);
    const ctx = buildLintContext(state.discover!, state.childDiscovers);
    state.lint = lintDocumentText(text, ctx);
    state.view = { kind: "tags" };
  });
}

/** Unique tag texts currently in the document (for highlight apply/clear). */
async function documentTagTexts(): Promise<string[]> {
  const text = await getDocumentText();
  return [...new Set(findTags(text).map((t) => t.tag))];
}

/**
 * Inserts a working sample layout for the current base object: a heading with
 * real fields, a repeating table over the first related list, and a totals
 * line — a one-click demonstration that also teaches the tag shapes.
 */
async function onInsertSampleLayout(): Promise<void> {
  const d = state.discover;
  if (!d || !state.api) return;
  await withBusy("Inserting sample layout…", async () => {
    const scalars = d.rootScalarMergeFields.slice(0, 3);
    const titleTag = scalars[0] ? scalarTag(scalars[0].key) : `{{${d.baseObjectApiName}.Name}}`;
    const detailLine = scalars
      .slice(1)
      .map((f) => `${f.label}: ${scalarTag(f.key, defaultFormatForType(f.type))}`)
      .join("   ");
    await insertParagraphBlock([
      `${d.baseObjectLabel}: ${titleTag}`,
      detailLine || `Generated {{Today}}`,
      `Prepared by {{RunningUser.Name}} on {{Today:MM/dd/yyyy}}`,
      "",
    ]);

    const rel = d.childRelationships.find((r) => r.relationshipName !== "Approvals");
    if (rel) {
      let child = state.childDiscovers.get(rel.childObjectApiName);
      if (!child) {
        try {
          child = await state.api!.discover(rel.childObjectApiName);
          state.childDiscovers.set(rel.childObjectApiName, child);
        } catch {
          child = undefined;
        }
      }
      if (child) {
        const cols = child.rootScalarMergeFields.slice(0, 3);
        if (cols.length > 0) {
          await insertLoopTable(
            rel.relationshipName,
            cols.map((f) => ({ inLoopKey: inLoopFieldKey(f.key), label: f.label })),
          );
          const numeric = child.rootScalarMergeFields.find((f) =>
            ["currency", "double", "int", "integer", "percent", "long"].includes(
              (f.type ?? "").toLowerCase(),
            ),
          );
          const totals = [aggregateTag("COUNT", rel.relationshipName) + ` ${rel.label} records`];
          if (numeric) {
            totals.push(
              `Total ${numeric.label}: ` +
                aggregateTag(
                  "SUM",
                  rel.relationshipName,
                  inLoopFieldKey(numeric.key),
                  defaultFormatForType(numeric.type) === "currency" ? "currency" : undefined,
                ),
            );
          }
          await insertParagraphBlock(totals);
        }
      }
    }
    state.notice = "Sample layout inserted — save it or tweak the tags to make it yours.";
    state.view = { kind: "main" };
  });
}

async function onToggleHighlights(on: boolean): Promise<void> {
  await withBusy(on ? "Highlighting tags…" : "Removing highlights…", async () => {
    await setTagHighlights(await documentTagTexts(), on ? "#FFF2CC" : null);
    state.highlightsOn = on;
  });
}

async function onPreview(recordId?: string): Promise<void> {
  if (!state.lastSavedVersionId) return;
  await withBusy("Generating preview…", async () => {
    const blob = await state.api!.preview({
      versionId: state.lastSavedVersionId!,
      ...(recordId?.trim() ? { recordId: recordId.trim() } : {}),
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "sliick-preview.docx";
    a.click();
    URL.revokeObjectURL(url);
    state.notice = "Preview downloaded — open it from your Downloads folder.";
  });
}

async function onConnect(): Promise<void> {
  await withBusy("Opening Salesforce login…", async () => {
    state.tokens = await login(state.settings.orgUrl, effectiveClientId(state.settings));
    state.api = new RealSliickClient(makeSession());
    state.view = { kind: "main" };
  });
  if (!state.error) await loadCatalog();
}

function onDisconnect(): void {
  clearTokens();
  state.tokens = null;
  state.api = null;
  state.view = { kind: "connect" };
  render();
}

// ---------------------------------------------------------------- render

function el(html: string): HTMLElement {
  const t = document.createElement("template");
  t.innerHTML = html.trim();
  return t.content.firstElementChild as HTMLElement;
}

function esc(s: string | null | undefined): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function render(): void {
  const app = document.getElementById("app")!;
  // Full re-render loses scroll position — capture and restore it.
  const scrollY = window.scrollY;
  app.innerHTML = "";
  app.appendChild(renderHeader());

  const pane = el(`<div class="pane"></div>`);
  if (state.busy) pane.appendChild(el(`<div class="banner info"><span class="spin">⟳</span> ${esc(state.busy)}</div>`));
  if (state.error) pane.appendChild(el(`<div class="banner error">${esc(state.error)}</div>`));
  if (state.notice) {
    pane.appendChild(el(`<div class="banner ok">${esc(state.notice)}</div>`));
    state.notice = null;
  }

  switch (state.view.kind) {
    case "settings":
      pane.appendChild(renderSettings());
      break;
    case "connect":
      pane.appendChild(renderConnect());
      break;
    case "loopWizard":
      pane.appendChild(renderLoopWizard(state.view));
      break;
    case "imageWizard":
      pane.appendChild(renderImageWizard());
      break;
    case "barcodeWizard":
      pane.appendChild(renderBarcodeWizard(state.view.initial));
      break;
    case "fieldOptions":
      pane.appendChild(renderFieldOptions(state.view.field, state.view.initial));
      break;
    case "editLoop":
      pane.appendChild(renderEditLoop(state.view));
      break;
    case "condWizard":
      pane.appendChild(renderCondWizard(state.view.mode));
      break;
    case "templates":
      pane.appendChild(renderTemplates(state.view.templates));
      break;
    case "tags":
      pane.appendChild(renderTagsPanel());
      break;
    case "save":
      pane.appendChild(renderSave());
      break;
    case "lintResults":
      pane.appendChild(renderLintResults(state.view.result));
      break;
    case "main":
      renderMain(pane);
      break;
  }
  app.appendChild(pane);

  const modeLabel = state.settings.mockMode
    ? "Demo mode (sample data — not connected to Salesforce)"
    : state.tokens
      ? `Connected: ${state.tokens.instanceUrl.replace(/^https:\/\//, "")}`
      : "Not connected";
  app.appendChild(el(`<div class="footer">${esc(modeLabel)}</div>`));
  window.scrollTo(0, scrollY);
}

function renderHeader(): HTMLElement {
  const header = el(`
    <div class="header">
      <span class="title">Sliick Docs</span>
      <span class="conn">${state.settings.mockMode ? "DEMO" : state.tokens ? "●" : "○"}</span>
      <button class="icon-btn" id="btn-settings" title="Settings">⚙</button>
    </div>`);
  header.querySelector("#btn-settings")!.addEventListener("click", () => {
    state.view = state.view.kind === "settings" ? { kind: "main" } : { kind: "settings" };
    render();
  });
  return header;
}

// ---------- settings ----------

function renderSettings(): HTMLElement {
  const s = state.settings;
  const root = el(`
    <div class="section">
      <div class="section-head">Settings</div>
      <div class="form">
        <label class="check">
          <input type="checkbox" id="set-mock" ${s.mockMode ? "checked" : ""} />
          Demo mode (sample data, no Salesforce connection)
        </label>
        <label>Salesforce org URL (My Domain)
          <input type="url" id="set-org" placeholder="https://yourdomain.my.salesforce.com"
                 value="${esc(s.orgUrl)}" ${s.mockMode ? "disabled" : ""} />
        </label>
        <label>External Client App consumer key
          <input type="text" id="set-client" placeholder="3MVG9…"
                 value="${esc(s.clientId)}" ${s.mockMode ? "disabled" : ""} />
          <span class="hint">Create a Local External Client App in your org and paste its consumer key here — see the setup guide.</span>
        </label>
        <div class="btn-row">
          <button class="btn primary" id="set-save">Save</button>
          <button class="btn secondary" id="set-cancel">Cancel</button>
          <span class="spacer"></span>
          ${state.tokens ? `<button class="btn secondary" id="set-disconnect">Sign out</button>` : ""}
        </div>
      </div>
    </div>`);

  const mockBox = root.querySelector<HTMLInputElement>("#set-mock")!;
  const orgInput = root.querySelector<HTMLInputElement>("#set-org")!;
  const clientInput = root.querySelector<HTMLInputElement>("#set-client")!;
  mockBox.addEventListener("change", () => {
    orgInput.disabled = mockBox.checked;
    clientInput.disabled = mockBox.checked;
  });

  root.querySelector("#set-save")!.addEventListener("click", () => {
    const mockMode = mockBox.checked;
    let orgUrl = state.settings.orgUrl;
    if (!mockMode) {
      const normalized = normalizeOrgUrl(orgInput.value);
      if (!normalized) {
        state.error = "Org URL must look like https://yourdomain.my.salesforce.com";
        render();
        return;
      }
      orgUrl = normalized;
    }
    // Blank consumer key is allowed — effectiveClientId() falls back to the
    // shared Sliick app key.
    state.settings = { mockMode, orgUrl, clientId: clientInput.value.trim() };
    saveSettings(state.settings);
    void initApi();
  });
  root.querySelector("#set-cancel")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  root.querySelector("#set-disconnect")?.addEventListener("click", onDisconnect);
  return root;
}

// ---------- connect ----------

function renderConnect(): HTMLElement {
  // Blank consumer key is fine — effectiveClientId() falls back to the shared
  // Sliick app key, so org URL alone is enough to connect.
  const configured = !!state.settings.orgUrl;
  const root = el(`
    <div class="section">
      <div class="section-head">Connect to Salesforce</div>
      <div class="form">
        ${
          configured
            ? `<p class="hint">Sign in to <b>${esc(state.settings.orgUrl)}</b> to load your org's merge fields.</p>
               <button class="btn primary" id="btn-login">Sign in to Salesforce</button>`
            : `<p class="hint">Set your org URL and External Client App consumer key in Settings (⚙) first — or switch on Demo mode to explore with sample data.</p>`
        }
      </div>
    </div>`);
  root.querySelector("#btn-login")?.addEventListener("click", () => void onConnect());
  return root;
}

// ---------- main: picker ----------

function renderMain(pane: HTMLElement): void {
  if (!state.api) {
    pane.appendChild(renderConnect());
    return;
  }

  // Base object select
  const objOptions = state.objects
    .map(
      (o) =>
        `<option value="${esc(o.apiName)}" ${o.apiName === state.baseObject ? "selected" : ""}>${esc(o.label)} (${esc(o.apiName)})</option>`,
    )
    .join("");
  const objSelect = el(`<select class="object-select" title="Template base object">${objOptions}</select>`) as HTMLSelectElement;
  objSelect.addEventListener("change", () => void selectBaseObject(objSelect.value));
  pane.appendChild(objSelect);

  // Search (Enter inserts the top match).
  const search = el(
    `<input class="search" type="text" placeholder="Search merge fields… (Enter inserts top match)" value="${esc(state.search)}" />`,
  ) as HTMLInputElement;
  search.addEventListener("input", () => {
    state.search = search.value;
    render();
    const again = document.querySelector<HTMLInputElement>(".search");
    if (again) {
      again.focus();
      again.setSelectionRange(again.value.length, again.value.length);
    }
  });
  search.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || !state.search.trim() || !state.discover) return;
    const all = [
      ...state.discover.rootScalarMergeFields,
      ...state.discover.parentLookupMergeFields,
      ...state.discover.runningUserMergeFields,
      ...state.discover.builtInMergeFields,
    ];
    const first = all.find(matchesSearch);
    if (first) void onInsertField(first);
  });
  pane.appendChild(search);

  const d = state.discover;
  if (!d) {
    if (!state.busy) pane.appendChild(el(`<div class="hint">Pick a base object to load fields.</div>`));
    return;
  }

  // One-click starter: a working sample layout for this object.
  const sampleRow = el(
    `<div class="btn-row"><button class="btn ghost" id="btn-sample" title="Insert a ready-made example with fields, a repeating table, and totals">✨ Insert sample layout</button></div>`,
  );
  sampleRow.querySelector("#btn-sample")!.addEventListener("click", () => void onInsertSampleLayout());
  pane.appendChild(sampleRow);

  // Selection smarts: offer to edit the tag under the cursor.
  if (state.editTag) {
    const card = el(`
      <div class="banner info tag-edit-card">
        <span class="lint-tag">${esc(truncateTag(state.editTag.tag))}</span>
        <button class="btn secondary" id="edit-tag-btn">Edit</button>
      </div>`);
    card.querySelector("#edit-tag-btn")!.addEventListener("click", () => {
      if (state.editTag) openTagEditor(state.editTag);
    });
    pane.appendChild(card);
  }

  // Selection smarts: cursor inside a loop → surface that list's fields first.
  if (state.cursorLoop) {
    const scope = scopeFieldsFor(state.cursorLoop);
    if (scope) {
      pane.appendChild(
        el(
          `<div class="banner info">Cursor is inside <b>{{#${esc(state.cursorLoop)}}}</b> — fields below insert in list scope.</div>`,
        ),
      );
      pane.appendChild(fieldSection(`${scope.label} (in this list)`, scope.fields));
    }
  }

  // Pinned + recent fields float to the top (per base object, localStorage).
  const allFields = [
    ...d.rootScalarMergeFields,
    ...d.parentLookupMergeFields,
    ...d.runningUserMergeFields,
    ...d.builtInMergeFields,
  ];
  const fieldByKey = new Map(allFields.map((f) => [f.key, f]));
  const pinnedFields = loadKeyList(pinsKey())
    .map((k) => fieldByKey.get(k))
    .filter((f): f is MergeFieldDef => !!f);
  const pinnedKeys = new Set(pinnedFields.map((f) => f.key));
  const recentFields = loadKeyList(recentsKey())
    .map((k) => fieldByKey.get(k))
    .filter((f): f is MergeFieldDef => !!f && !pinnedKeys.has(f.key));
  if (pinnedFields.length > 0) pane.appendChild(fieldSection("📌 Pinned", pinnedFields));
  if (recentFields.length > 0) pane.appendChild(fieldSection("Recent", recentFields));

  pane.appendChild(fieldSection("Fields", d.rootScalarMergeFields));
  pane.appendChild(fieldSection("Parent fields", d.parentLookupMergeFields));
  pane.appendChild(fieldSection("Running user", d.runningUserMergeFields));
  pane.appendChild(fieldSection("Built-ins", d.builtInMergeFields));
  pane.appendChild(relationshipSection(d.childRelationships));

  // Logic + save actions
  const caps = state.capabilities?.features;
  const actions = el(`<div class="btn-row"></div>`);
  if (caps?.conditionals) {
    const b = el(`<button class="btn secondary">+ Condition</button>`);
    b.addEventListener("click", () => {
      state.view = { kind: "condWizard", mode: "if" };
      render();
    });
    actions.appendChild(b);
  }
  if (caps?.inverseConditionals) {
    const b = el(`<button class="btn secondary" title="Show content only when a field is blank or false">+ If blank</button>`);
    b.addEventListener("click", () => {
      state.view = { kind: "condWizard", mode: "inverse" };
      render();
    });
    actions.appendChild(b);
  }
  if (caps?.imageFields) {
    const b = el(`<button class="btn secondary" title="Insert an image from a file/field on the record">+ Image</button>`);
    b.addEventListener("click", () => {
      state.view = { kind: "imageWizard" };
      render();
    });
    actions.appendChild(b);
  }
  if (caps?.barcodes) {
    const b = el(`<button class="btn secondary" title="Insert a barcode or QR code encoding a field's value">+ Barcode</button>`);
    b.addEventListener("click", () => {
      state.view = { kind: "barcodeWizard" };
      render();
    });
    actions.appendChild(b);
  }
  actions.appendChild(el(`<span class="spacer"></span>`));
  const tagsBtn = el(`<button class="btn secondary" title="Check every merge tag in the document and jump to problems">Review tags</button>`);
  tagsBtn.addEventListener("click", () => void onReviewTags());
  actions.appendChild(tagsBtn);
  const libBtn = el(`<button class="btn secondary" title="Browse templates saved in Salesforce">My templates</button>`);
  libBtn.addEventListener("click", () => void onTemplatesOpen());
  actions.appendChild(libBtn);
  const saveBtn = el(`<button class="btn primary">Save to Salesforce</button>`);
  saveBtn.addEventListener("click", () => {
    state.view = { kind: "save" };
    render();
  });
  actions.appendChild(saveBtn);
  pane.appendChild(actions);
}

function matchesSearch(f: MergeFieldDef): boolean {
  const q = state.search.trim().toLowerCase();
  if (!q) return true;
  return f.label.toLowerCase().includes(q) || f.key.toLowerCase().includes(q);
}

function fieldSection(title: string, fields: MergeFieldDef[]): HTMLElement {
  const visible = fields.filter(matchesSearch);
  const collapsed = state.collapsed.has(title) && !state.search.trim();
  const root = el(`
    <div class="section ${collapsed ? "collapsed" : ""}">
      <div class="section-head">${esc(title)} <span class="chev">${collapsed ? "▸" : "▾"} ${visible.length}</span></div>
      <div class="section-body"></div>
    </div>`);
  root.querySelector(".section-head")!.addEventListener("click", () => {
    if (state.collapsed.has(title)) state.collapsed.delete(title);
    else state.collapsed.add(title);
    render();
  });
  const body = root.querySelector(".section-body")!;
  if (visible.length === 0) {
    body.appendChild(el(`<div class="hint">No matching fields.</div>`));
  }
  const optionsOn =
    state.capabilities?.features.fallbackText === true ||
    state.capabilities?.features.localeFormats === true;
  for (const f of visible) {
    const row = el(`
      <div class="field-row-wrap">
        <button class="field-row" title="Insert ${esc(scalarTag(f.key, defaultFormatForType(f.type)))}">
          <span class="f-label">${esc(f.label)}</span>
          <span class="f-type">${esc(f.type)}</span>
        </button>
        ${optionsOn ? `<button class="row-options" title="Insert with options (format, locale, fallback)">⋯</button>` : ""}
      </div>`);
    row.querySelector(".field-row")!.addEventListener("click", () => void onInsertField(f));
    row.querySelector(".row-options")?.addEventListener("click", () => {
      state.view = { kind: "fieldOptions", field: f };
      render();
    });
    body.appendChild(row);
  }
  return root;
}

function relationshipSection(rels: ChildRelationshipDef[]): HTMLElement {
  const canLoop = state.capabilities?.features.childLoops !== false;
  const root = el(`
    <div class="section">
      <div class="section-head">Related lists (repeating tables) <span class="chev">${rels.length}</span></div>
      <div class="section-body"></div>
    </div>`);
  const body = root.querySelector(".section-body")!;
  if (!canLoop) {
    body.appendChild(el(`<div class="hint">Repeating tables aren't supported by the installed Sliick Docs version.</div>`));
    return root;
  }
  if (rels.length === 0) {
    body.appendChild(el(`<div class="hint">No child relationships on this object.</div>`));
  }
  for (const rel of rels) {
    const row = el(`
      <button class="field-row" title="Insert a repeating table over ${esc(rel.relationshipName)}">
        <span class="f-label">${esc(rel.label)}</span>
        <span class="f-type">table</span>
      </button>`);
    row.addEventListener("click", () => void onLoopWizardOpen(rel));
    body.appendChild(row);
  }
  return root;
}

// ---------- loop wizard ----------

/** Currency/percent format for an aggregated numeric field, else undefined. */
function aggregateFormatFor(inLoopKey: string, numericFields: MergeFieldDef[]): string | undefined {
  const match = numericFields.find((f) => inLoopFieldKey(f.key) === inLoopKey);
  if (!match) return undefined;
  const fmt = defaultFormatForType(match.type);
  return fmt === "currency" || fmt === "percent" ? fmt : undefined;
}

function renderLoopWizard(view: Extract<View, { kind: "loopWizard" }>): HTMLElement {
  const { rel, childFields, childRels, childChecked, nestedRel, nestedFields, nestedChecked } =
    view;
  const checks = childFields
    .map(
      (f, i) =>
        `<label><input type="checkbox" data-i="${i}" ${childChecked.includes(i) ? "checked" : ""}/> ${esc(f.label)} <span class="f-type">${esc(f.type)}</span></label>`,
    )
    .join("");
  const numericFields = childFields.filter((f) =>
    ["currency", "double", "int", "integer", "percent", "long"].includes((f.type ?? "").toLowerCase()),
  );
  const aggFieldOptions = numericFields
    .map((f) => `<option value="${esc(inLoopFieldKey(f.key))}">${esc(f.label)}</option>`)
    .join("");
  // No totals over the synthetic Approvals relationship — the aggregate fetch
  // plan only queries real child relationships.
  const aggregatesOn =
    state.capabilities?.features.aggregates !== false && rel.relationshipName !== "Approvals";
  const nestedOn = state.capabilities?.features.nestedLoops === true && childRels.length > 0;
  const filtersOn = state.capabilities?.features.loopFilters === true;
  const modFieldOptions = childFields
    .map(
      (f) =>
        `<option value="${esc(inLoopFieldKey(f.key))}" data-type="${esc(f.type)}">${esc(f.label)}</option>`,
    )
    .join("");
  const root = el(`
    <div class="section">
      <div class="section-head">Repeating table — ${esc(rel.label)}</div>
      <div class="form">
        <p class="hint">Pick the columns. A table is inserted where your cursor is:
        a header row plus one row that repeats for every ${esc(rel.label)} record.</p>
        <div class="col-list" id="lw-cols">${checks}</div>
        ${
          filtersOn
            ? `<p class="hint">Filter and sort (optional):</p>
        <label>Only include records where
          <select id="lw-f-field"><option value="">— no filter —</option>${modFieldOptions}</select>
        </label>
        <div class="btn-row">
          <select id="lw-f-op">
            <option value="=">=</option><option value="!=">≠</option>
            <option value=">">&gt;</option><option value="<">&lt;</option>
            <option value=">=">≥</option><option value="<=">≤</option>
            <option value="contains">contains</option>
          </select>
          <input type="text" id="lw-f-value" placeholder="value" style="flex:1" />
        </div>
        <label class="check"><input type="checkbox" id="lw-f-not" /> NOT — exclude matching records</label>
        <div class="btn-row">
          <label style="flex:1">Sort by
            <select id="lw-sort"><option value="">— document order —</option>${modFieldOptions}</select>
          </label>
          <select id="lw-dir">
            <option value="">Ascending</option>
            <option value="DESC">Descending</option>
          </select>
        </div>`
            : ""
        }
        <div class="btn-row">
          <button class="btn primary" id="lw-insert">Insert table</button>
          <button class="btn secondary" id="lw-cancel">Cancel</button>
        </div>
        ${
          aggregatesOn
            ? `<hr/>
        <p class="hint">Insert a total (placed outside the table, e.g. a summary row):</p>
        <div class="btn-row">
          <select id="lw-agg-fn">
            <option value="SUM">Sum</option><option value="COUNT">Count</option>
            <option value="AVG">Average</option><option value="MIN">Min</option><option value="MAX">Max</option>
          </select>
          <select id="lw-agg-field">${aggFieldOptions}</select>
          <button class="btn secondary" id="lw-agg-insert">Insert total</button>
        </div>`
            : ""
        }
        ${
          nestedOn
            ? `<hr/>
        <p class="hint">Or insert a nested list: the checked columns above for each
        ${esc(rel.label)} record, then a related list inside it (one nesting level).</p>
        <label>Nested related list
          <select id="lw-nest-rel">
            <option value="">— none —</option>
            ${childRels
              .map(
                (r, i) =>
                  `<option value="${i}" ${nestedRel?.relationshipName === r.relationshipName ? "selected" : ""}>${esc(r.label)}</option>`,
              )
              .join("")}
          </select>
        </label>
        ${
          nestedRel && nestedFields
            ? `<div class="col-list" id="lw-nest-cols">${nestedFields
                .map(
                  (f, i) =>
                    `<label><input type="checkbox" data-i="${i}" ${nestedChecked?.includes(i) ? "checked" : ""}/> ${esc(f.label)} <span class="f-type">${esc(f.type)}</span></label>`,
                )
                .join("")}</div>
        <div class="btn-row">
          <button class="btn primary" id="lw-nest-insert">Insert nested list</button>
        </div>`
            : ""
        }`
            : ""
        }
      </div>
    </div>`);

  const aggInsert = root.querySelector("#lw-agg-insert");
  if (aggInsert) {
    aggInsert.addEventListener("click", () => {
      const fn = root.querySelector<HTMLSelectElement>("#lw-agg-fn")!.value as AggregateFn;
      const fieldKey = root.querySelector<HTMLSelectElement>("#lw-agg-field")!.value;
      if (fn !== "COUNT" && !fieldKey) {
        state.error = "Pick a numeric field for this total (or choose Count).";
        render();
        return;
      }
      const format = fn !== "COUNT" ? aggregateFormatFor(fieldKey, numericFields) : undefined;
      void withBusy("Inserting total…", async () => {
        await insertScalar(aggregateTag(fn, rel.relationshipName, fieldKey || undefined, format));
        state.view = { kind: "main" };
      });
    });
  }

  // Checked columns live in the view state (not the DOM) so re-renders —
  // e.g. after the nested-list discover() — don't reset the user's picks.
  // The change listeners record without re-rendering; the DOM already shows
  // the new state.
  const bindChecks = (containerId: string, checked: number[]): void => {
    root
      .querySelectorAll<HTMLInputElement>(`#${containerId} input[type=checkbox]`)
      .forEach((box) => {
        box.addEventListener("change", () => {
          const i = Number(box.dataset.i);
          const at = checked.indexOf(i);
          if (box.checked && at === -1) checked.push(i);
          if (!box.checked && at !== -1) checked.splice(at, 1);
        });
      });
  };
  bindChecks("lw-cols", childChecked);
  if (nestedChecked) bindChecks("lw-nest-cols", nestedChecked);

  // Column order follows field order regardless of check order.
  const pickedFields = (checked: number[], fields: MergeFieldDef[]): MergeFieldDef[] =>
    [...checked]
      .sort((a, b) => a - b)
      .map((i) => fields[i])
      .filter((f): f is MergeFieldDef => !!f);

  // grammar-v2: optional WHERE clause + ORDER BY read from the filter section.
  const NUMERIC_FILTER_TYPES = new Set(["currency", "double", "int", "integer", "percent", "long"]);
  const readLoopModifiers = (): LoopModifiers | undefined => {
    if (!filtersOn) return undefined;
    const mods: LoopModifiers = {};
    const filterKey = root.querySelector<HTMLSelectElement>("#lw-f-field")?.value ?? "";
    const filterValue = root.querySelector<HTMLInputElement>("#lw-f-value")?.value.trim() ?? "";
    if (filterKey && filterValue) {
      const fieldType = (
        childFields.find((f) => inLoopFieldKey(f.key) === filterKey)?.type ?? "string"
      ).toLowerCase();
      mods.where = conditionExpressionText(
        [
          {
            fieldKey: `${rel.relationshipName}.${filterKey}`,
            operator: root.querySelector<HTMLSelectElement>("#lw-f-op")!.value as ConditionOperator,
            value: filterValue,
            quoteValue: !NUMERIC_FILTER_TYPES.has(fieldType),
            negate: root.querySelector<HTMLInputElement>("#lw-f-not")!.checked,
          },
        ],
        "AND",
      );
    }
    const sortKey = root.querySelector<HTMLSelectElement>("#lw-sort")?.value ?? "";
    if (sortKey) {
      mods.orderBy = sortKey;
      mods.descending = (root.querySelector<HTMLSelectElement>("#lw-dir")?.value ?? "") === "DESC";
    }
    return mods.where || mods.orderBy ? mods : undefined;
  };

  root.querySelector("#lw-insert")!.addEventListener("click", () => {
    const picked = pickedFields(childChecked, childFields);
    if (picked.length === 0) {
      state.error = "Pick at least one column.";
      render();
      return;
    }
    void onLoopInsert(rel, picked, readLoopModifiers());
  });

  root.querySelector("#lw-nest-rel")?.addEventListener("change", (e) => {
    const idx = (e.target as HTMLSelectElement).value;
    void onNestedRelPick(idx === "" ? undefined : childRels[Number(idx)]);
  });
  root.querySelector("#lw-nest-insert")?.addEventListener("click", () => {
    if (!nestedRel || !nestedFields || !nestedChecked) return;
    const childPicked = pickedFields(childChecked, childFields);
    const nestedPicked = pickedFields(nestedChecked, nestedFields);
    if (childPicked.length === 0 || nestedPicked.length === 0) {
      state.error = "Pick at least one field at each level.";
      render();
      return;
    }
    void onNestedLoopInsert(rel, childPicked, nestedRel, nestedPicked, readLoopModifiers());
  });

  root.querySelector("#lw-cancel")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- conditional wizard ----------

function renderCondWizard(mode: "if" | "inverse"): HTMLElement {
  const fields = state.discover?.rootScalarMergeFields ?? [];
  const fieldOptions = fields
    .map((f) => `<option value="${esc(f.key)}" data-type="${esc(f.type)}">${esc(f.label)}</option>`)
    .join("");
  const isIf = mode === "if";
  const compoundOn = state.capabilities?.features.compoundConditions !== false;
  const numericTypes = new Set(["currency", "int", "integer", "double", "long", "percent"]);
  const typeOf = (key: string): string => fields.find((f) => f.key === key)?.type ?? "string";
  const root = el(`
    <div class="section">
      <div class="section-head">${isIf ? "Conditional content" : "Show when blank"}</div>
      <div class="form">
        <p class="hint">${
          isIf
            ? "Content between the tags shows only when the condition is true. Select text first to wrap it, or insert fresh tags at the cursor."
            : "Content between the tags shows only when the field is blank or false."
        }</p>
        <label>Field <select id="cw-field">${fieldOptions}</select></label>
        ${
          isIf
            ? `<label>Operator
                 <select id="cw-op">
                   <option value="=">=</option><option value="!=">≠</option>
                   <option value=">">&gt;</option><option value="<">&lt;</option>
                   <option value=">=">≥</option><option value="<=">≤</option>
                   <option value="contains">contains</option>
                 </select>
               </label>
               <label>Value <input type="text" id="cw-value" placeholder="e.g. 50000 or Closed Won" /></label>
               <label class="check"><input type="checkbox" id="cw-not" /> NOT — show when this is false</label>
               ${
                 compoundOn
                   ? `<label>Add another condition
                        <select id="cw-conn">
                          <option value="">— none —</option>
                          <option value="AND">AND</option>
                          <option value="OR">OR</option>
                        </select>
                      </label>
                      <label>Field 2 <select id="cw-field2"><option value=""></option>${fieldOptions}</select></label>
                      <label>Operator 2
                        <select id="cw-op2">
                          <option value="=">=</option><option value="!=">≠</option>
                          <option value=">">&gt;</option><option value="<">&lt;</option>
                          <option value=">=">≥</option><option value="<=">≤</option>
                          <option value="contains">contains</option>
                        </select>
                      </label>
                      <label>Value 2 <input type="text" id="cw-value2" placeholder="optional second value" /></label>
                      <label class="check"><input type="checkbox" id="cw-not2" /> NOT — invert the second condition</label>`
                   : ""
               }
               <label class="check"><input type="checkbox" id="cw-else" /> Include an otherwise (else) branch</label>`
            : ""
        }
        <div class="btn-row">
          <button class="btn primary" id="cw-insert">Insert</button>
          <button class="btn secondary" id="cw-cancel">Cancel</button>
        </div>
      </div>
    </div>`);

  root.querySelector("#cw-insert")!.addEventListener("click", () => {
    const fieldSel = root.querySelector<HTMLSelectElement>("#cw-field")!;
    const fieldKey = fieldSel.value;
    if (!fieldKey) return;
    if (!isIf) {
      void withBusy("Inserting…", async () => {
        await insertInverse(fieldKey);
        state.view = { kind: "main" };
      });
      return;
    }
    const op = root.querySelector<HTMLSelectElement>("#cw-op")!.value as ConditionOperator;
    const value = root.querySelector<HTMLInputElement>("#cw-value")!.value.trim();
    if (!value) {
      state.error = "Enter a value to compare against.";
      render();
      return;
    }
    const withElse = root.querySelector<HTMLInputElement>("#cw-else")!.checked;
    const clause1: ConditionClause = {
      fieldKey,
      operator: op,
      value,
      quoteValue: !numericTypes.has(typeOf(fieldKey)),
      negate: root.querySelector<HTMLInputElement>("#cw-not")!.checked,
    };

    // Optional second clause → compound AND/OR condition.
    const connector = root.querySelector<HTMLSelectElement>("#cw-conn")?.value;
    const field2 = root.querySelector<HTMLSelectElement>("#cw-field2")?.value ?? "";
    const value2 = root.querySelector<HTMLInputElement>("#cw-value2")?.value.trim() ?? "";
    if (connector && field2 && value2) {
      const op2 = root.querySelector<HTMLSelectElement>("#cw-op2")!.value as ConditionOperator;
      const clause2: ConditionClause = {
        fieldKey: field2,
        operator: op2,
        value: value2,
        quoteValue: !numericTypes.has(typeOf(field2)),
        negate: root.querySelector<HTMLInputElement>("#cw-not2")!.checked,
      };
      const { open, elseTag, close } = compoundConditionTags(
        [clause1, clause2],
        connector as "AND" | "OR",
        withElse,
      );
      void withBusy("Inserting…", async () => {
        await insertConditionalTags(open, elseTag, close);
        state.view = { kind: "main" };
      });
      return;
    }

    void withBusy("Inserting…", async () => {
      await insertConditional({ ...clause1, withElse });
      state.view = { kind: "main" };
    });
  });
  root.querySelector("#cw-cancel")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- image wizard ----------

function renderImageWizard(): HTMLElement {
  const fields = state.discover?.rootScalarMergeFields ?? [];
  const fieldOptions = fields
    .map((f) => `<option value="${esc(f.key)}">${esc(f.label)}</option>`)
    .join("");
  const root = el(`
    <div class="section">
      <div class="section-head">Insert image</div>
      <div class="form">
        <p class="hint">Pick a field whose value is a Salesforce File (ContentVersion/ContentDocument Id).
        The image is embedded into the document at generation time.</p>
        <label>Image field <select id="iw-field">${fieldOptions}</select></label>
        <div class="btn-row">
          <label>Width px <input type="text" id="iw-w" placeholder="200" style="width:70px" /></label>
          <label>Height px <input type="text" id="iw-h" placeholder="120" style="width:70px" /></label>
        </div>
        <div class="btn-row">
          <button class="btn primary" id="iw-insert">Insert image</button>
          <button class="btn secondary" id="iw-cancel">Cancel</button>
        </div>
      </div>
    </div>`);
  root.querySelector("#iw-insert")!.addEventListener("click", () => {
    const fieldKey = root.querySelector<HTMLSelectElement>("#iw-field")!.value;
    if (!fieldKey) return;
    const w = Number(root.querySelector<HTMLInputElement>("#iw-w")!.value) || undefined;
    const h = Number(root.querySelector<HTMLInputElement>("#iw-h")!.value) || undefined;
    void withBusy("Inserting image…", async () => {
      await insertScalar(imageTag(fieldKey, w, h));
      state.view = { kind: "main" };
    });
  });
  root.querySelector("#iw-cancel")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- template library ----------

function renderTemplates(templates: TemplateSummary[]): HTMLElement {
  const root = el(`
    <div class="section">
      <div class="section-head">My templates${state.baseObject ? ` — ${esc(state.baseObject)}` : ""}</div>
      <div class="form">
        <p class="hint">Uploaded templates saved in Salesforce for this base object.
        Pick one to save the current document as its next version.</p>
        <div class="section-body" id="tpl-list"></div>
        <div class="btn-row">
          <button class="btn secondary" id="tpl-back">Back to fields</button>
        </div>
      </div>
    </div>`);
  const list = root.querySelector("#tpl-list")!;
  if (templates.length === 0) {
    list.appendChild(el(`<div class="hint">No templates yet — Save to Salesforce creates the first one.</div>`));
  }
  for (const t of templates) {
    const date = t.lastModifiedDate ? new Date(t.lastModifiedDate).toLocaleDateString() : "";
    const row = el(`
      <button class="field-row" title="Save the current document as a new version of ${esc(t.name)}">
        <span class="f-label">${esc(t.name)}</span>
        <span class="f-type"><span class="lint-status ${t.validationStatus}">${t.validationStatus}</span> ${esc(date)}</span>
      </button>`);
    row.addEventListener("click", () => {
      state.activeTemplate = { templateId: t.templateId, name: t.name };
      state.view = { kind: "save" };
      render();
    });
    list.appendChild(row);
  }
  root.querySelector("#tpl-back")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- tags panel (local lint + navigator) ----------

function renderTagsPanel(): HTMLElement {
  const lint = state.lint;
  const banner = !lint
    ? ""
    : lint.entries.length === 0
      ? `<div class="banner info">No merge tags in the document yet.</div>`
      : lint.hasProblems
        ? `<div class="banner error">Some tags need attention — click one to jump to it.</div>`
        : `<div class="banner ok">All ${lint.entries.length} tags check out. The full validation still runs on save.</div>`;
  const errorRows = (lint?.errors ?? [])
    .map((e) => `<div class="banner error">${esc(e)}</div>`)
    .join("");
  const items = (lint?.entries ?? [])
    .map((t, i) => {
      const detail = t.suggestion
        ? `<div class="lint-suggest">Did you mean <b>${esc(t.suggestion)}</b>?</div>`
        : t.note
          ? `<div class="lint-suggest">${esc(t.note)}</div>`
          : "";
      return `<li data-i="${i}" class="tag-row" title="Jump to this tag in the document">
        <div><span class="lint-tag">${esc(t.tag)}</span>${detail}</div>
        <span class="lint-status ${t.status}">${t.status}</span></li>`;
    })
    .join("");
  const root = el(`
    <div class="section">
      <div class="section-head">Tags in this document</div>
      <div class="form">
        ${banner}
        ${errorRows}
        ${items ? `<ul class="lint-list">${items}</ul>` : ""}
        <label class="check">
          <input type="checkbox" id="tp-highlight" ${state.highlightsOn ? "checked" : ""} />
          Highlight tags in the document (removed automatically on save)
        </label>
        <div class="btn-row">
          <button class="btn secondary" id="tp-back">Back to fields</button>
          <span class="spacer"></span>
          <button class="btn primary" id="tp-refresh">Re-check</button>
        </div>
      </div>
    </div>`);
  root.querySelectorAll<HTMLElement>(".tag-row").forEach((row) => {
    row.addEventListener("click", () => {
      const entry = state.lint?.entries[Number(row.dataset.i)];
      if (entry) {
        void withBusy(null, async () => {
          await selectTagOccurrence(entry.tag, entry.occurrence);
        });
      }
    });
  });
  root.querySelector("#tp-highlight")!.addEventListener("change", (e) => {
    void onToggleHighlights((e.target as HTMLInputElement).checked);
  });
  root.querySelector("#tp-back")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  root.querySelector("#tp-refresh")!.addEventListener("click", () => void onReviewTags());
  return root;
}

// ---------- barcode wizard ----------

/** Shortens long tags for the edit card. */
function truncateTag(tag: string): string {
  return tag.length > 44 ? `${tag.slice(0, 41)}…` : tag;
}

function renderBarcodeWizard(
  initial?: { key: string; barcodeType: "code128" | "qr"; size?: string },
): HTMLElement {
  const fields = state.discover?.rootScalarMergeFields ?? [];
  const knownKey = fields.some((f) => f.key === initial?.key);
  const fieldOptions =
    (initial && !knownKey
      ? `<option value="${esc(initial.key)}" selected>${esc(initial.key)}</option>`
      : "") +
    fields
      .map(
        (f) =>
          `<option value="${esc(f.key)}" ${f.key === initial?.key ? "selected" : ""}>${esc(f.label)}</option>`,
      )
      .join("");
  const root = el(`
    <div class="section">
      <div class="section-head">Insert barcode / QR code</div>
      <div class="form">
        <p class="hint">The field's value (e.g. an invoice number or URL) is encoded
        into the barcode when the document is generated.</p>
        <label>Field <select id="bw-field">${fieldOptions}</select></label>
        <label>Type
          <select id="bw-type">
            <option value="code128" ${initial?.barcodeType !== "qr" ? "selected" : ""}>Barcode (Code 128)</option>
            <option value="qr" ${initial?.barcodeType === "qr" ? "selected" : ""}>QR code</option>
          </select>
        </label>
        <div class="btn-row">
          <label>Size px <input type="text" id="bw-size" placeholder="auto" style="width:90px" value="${esc(initial?.size ?? "")}" /></label>
          <span class="hint" id="bw-size-hint">width×height, e.g. 250x80</span>
        </div>
        <div class="btn-row">
          <button class="btn primary" id="bw-insert">${initial ? "Update" : "Insert"}</button>
          <button class="btn secondary" id="bw-cancel">Cancel</button>
        </div>
      </div>
    </div>`);
  const typeSel = root.querySelector<HTMLSelectElement>("#bw-type")!;
  const sizeHint = root.querySelector<HTMLElement>("#bw-size-hint")!;
  typeSel.addEventListener("change", () => {
    sizeHint.textContent =
      typeSel.value === "qr" ? "square size, e.g. 150" : "width×height, e.g. 250x80";
  });
  root.querySelector("#bw-insert")!.addEventListener("click", () => {
    const fieldKey = root.querySelector<HTMLSelectElement>("#bw-field")!.value;
    if (!fieldKey) return;
    const barType = typeSel.value as "code128" | "qr";
    const size = root.querySelector<HTMLInputElement>("#bw-size")!.value.trim().toLowerCase();
    if (size && !/^\d+(x\d+)?$/.test(size)) {
      state.error = "Size must be a number (QR) or widthxheight (barcode), e.g. 150 or 250x80.";
      render();
      return;
    }
    void withBusy("Inserting barcode…", async () => {
      await insertOrReplaceScalar(barcodeTag(fieldKey, barType, size || undefined));
      state.view = { kind: "main" };
    });
  });
  root.querySelector("#bw-cancel")!.addEventListener("click", () => {
    state.replaceTarget = null;
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- edit an existing loop's filter/sort ----------

function renderEditLoop(view: Extract<View, { kind: "editLoop" }>): HTMLElement {
  const scope = scopeFieldsFor(view.relationship);
  const sortOptions =
    `<option value="">— document order —</option>` +
    (scope?.fields ?? [])
      .map(
        (f) =>
          `<option value="${esc(f.key)}" ${f.key === view.orderBy ? "selected" : ""}>${esc(f.label)}</option>`,
      )
      .join("") +
    (view.orderBy && !(scope?.fields ?? []).some((f) => f.key === view.orderBy)
      ? `<option value="${esc(view.orderBy)}" selected>${esc(view.orderBy)}</option>`
      : "");
  const root = el(`
    <div class="section">
      <div class="section-head">Repeating list — ${esc(view.relationship)}</div>
      <div class="form">
        <p class="hint">Update this list's filter and sort. The filter uses the
        condition grammar with the "${esc(view.relationship)}." prefix, e.g.
        ${esc(view.relationship)}.Amount &gt; 100 AND ${esc(view.relationship)}.Status = 'Open'.</p>
        <label>Filter (blank = all records)
          <input type="text" id="el-where" value="${esc(view.where ?? "")}"
                 placeholder="${esc(view.relationship)}.Amount > 100" />
        </label>
        <div class="btn-row">
          <label style="flex:1">Sort by <select id="el-sort">${sortOptions}</select></label>
          <select id="el-dir">
            <option value="" ${view.descending ? "" : "selected"}>Ascending</option>
            <option value="DESC" ${view.descending ? "selected" : ""}>Descending</option>
          </select>
        </div>
        <div class="btn-row">
          <button class="btn primary" id="el-update">Update</button>
          <button class="btn secondary" id="el-cancel">Cancel</button>
        </div>
      </div>
    </div>`);
  root.querySelector("#el-update")!.addEventListener("click", () => {
    const where = root.querySelector<HTMLInputElement>("#el-where")!.value.trim();
    const orderBy = root.querySelector<HTMLSelectElement>("#el-sort")!.value;
    const descending = root.querySelector<HTMLSelectElement>("#el-dir")!.value === "DESC";
    const open = loopTags(view.relationship, {
      where: where || undefined,
      orderBy: orderBy || undefined,
      descending,
    }).open;
    void withBusy("Updating list…", async () => {
      await insertOrReplaceScalar(open);
      state.view = { kind: "main" };
    });
  });
  root.querySelector("#el-cancel")!.addEventListener("click", () => {
    state.replaceTarget = null;
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- field insert-with-options ----------

/** Locale choices mirroring TemplateFormatterService's override table. */
const LOCALE_CHOICES: [string, string][] = [
  ["en_US", "English (US)"],
  ["en_GB", "English (UK)"],
  ["en_AU", "English (Australia)"],
  ["en_CA", "English (Canada)"],
  ["de_DE", "German (Germany)"],
  ["de_AT", "German (Austria)"],
  ["de_CH", "German (Switzerland)"],
  ["fr_FR", "French (France)"],
  ["fr_CA", "French (Canada)"],
  ["es_ES", "Spanish (Spain)"],
  ["es_MX", "Spanish (Mexico)"],
  ["it_IT", "Italian"],
  ["nl_NL", "Dutch"],
  ["pt_BR", "Portuguese (Brazil)"],
  ["pt_PT", "Portuguese (Portugal)"],
  ["ja_JP", "Japanese"],
  ["zh_CN", "Chinese (Simplified)"],
  ["ko_KR", "Korean"],
  ["sv_SE", "Swedish"],
  ["da_DK", "Danish"],
  ["nb_NO", "Norwegian"],
  ["fi_FI", "Finnish"],
  ["pl_PL", "Polish"],
  ["ru_RU", "Russian"],
];

/** Format choices per field type: [suffix ('' = raw), label]. */
function formatChoicesForType(sfType?: string): [string, string][] {
  switch ((sfType ?? "").toLowerCase()) {
    case "currency":
      return [["currency", "Currency ($1,234.50)"], ["number", "Number (1,234.5)"], ["", "Raw value"]];
    case "percent":
      return [["percent", "Percent (75%)"], ["number", "Number"], ["", "Raw value"]];
    case "double":
    case "int":
    case "integer":
    case "long":
      return [["number", "Number (1,234.5)"], ["", "Raw value"]];
    case "date":
      return [["MM/dd/yyyy", "03/18/2026"], ["date", "Your locale's date"], ["yyyy-MM-dd", "2026-03-18"], ["", "Raw value"]];
    case "datetime":
      return [["MM/dd/yyyy h:mm a", "03/18/2026 2:30 PM"], ["datetime", "Your locale's date & time"], ["", "Raw value"]];
    case "boolean":
      return [["checkbox", "Checkbox [X] / [ ]"], ["", "true / false"]];
    case "picklist":
    case "multipicklist":
      return [["label", "Display label"], ["", "Stored API value"]];
    default:
      return [["", "Raw value"]];
  }
}

function renderFieldOptions(
  field: MergeFieldDef,
  initial?: { format?: string; locale?: string; fallback?: string },
): HTMLElement {
  const caps = state.capabilities?.features;
  const fallbackOn = caps?.fallbackText === true;
  const localeOn = caps?.localeFormats === true;
  const choices = formatChoicesForType(field.type);
  const selectedFormat = initial ? (initial.format ?? "") : (defaultFormatForType(field.type) ?? "");
  const formatOptions =
    (selectedFormat && !choices.some(([value]) => value === selectedFormat)
      ? `<option value="${esc(selectedFormat)}" selected>Keep current (${esc(selectedFormat)})</option>`
      : "") +
    choices
      .map(
        ([value, label]) =>
          `<option value="${esc(value)}" ${value === selectedFormat ? "selected" : ""}>${esc(label)}</option>`,
      )
      .join("");
  const localeOptions =
    `<option value="">Your locale (default)</option>` +
    LOCALE_CHOICES.map(
      ([value, label]) =>
        `<option value="${esc(value)}" ${value === initial?.locale ? "selected" : ""}>${esc(label)}</option>`,
    ).join("");
  const root = el(`
    <div class="section">
      <div class="section-head">${initial ? "Edit" : "Insert"} ${esc(field.label)}</div>
      <div class="form">
        <label>Format <select id="fo-format">${formatOptions}</select></label>
        ${
          localeOn
            ? `<label>Locale <select id="fo-locale" title="Formats numbers and dates for a specific locale">${localeOptions}</select></label>`
            : ""
        }
        ${
          fallbackOn
            ? `<label>If blank, show
                 <input type="text" id="fo-fallback" placeholder="e.g. N/A (optional)" value="${esc(initial?.fallback ?? "")}" /></label>`
            : ""
        }
        <label class="check">
          <input type="checkbox" id="fo-pin" ${isPinned(field.key) ? "checked" : ""} />
          📌 Pin this field to the top of the list
        </label>
        <div class="btn-row">
          <button class="btn primary" id="fo-insert">${initial ? "Update" : "Insert"}</button>
          <button class="btn secondary" id="fo-cancel">Cancel</button>
        </div>
      </div>
    </div>`);
  root.querySelector("#fo-pin")!.addEventListener("change", () => togglePin(field.key));
  const formatSel = root.querySelector<HTMLSelectElement>("#fo-format")!;
  const localeSel = root.querySelector<HTMLSelectElement>("#fo-locale");
  const syncLocaleEnabled = (): void => {
    if (localeSel) localeSel.disabled = !LOCALE_KEYWORD_FORMATS.has(formatSel.value);
  };
  formatSel.addEventListener("change", syncLocaleEnabled);
  syncLocaleEnabled();
  root.querySelector("#fo-insert")!.addEventListener("click", () => {
    const format = formatSel.value || undefined;
    const locale = localeSel?.disabled ? undefined : localeSel?.value || undefined;
    const fallback = root.querySelector<HTMLInputElement>("#fo-fallback")?.value ?? "";
    noteRecentField(field.key);
    void withBusy(null, async () => {
      await insertOrReplaceScalar(scalarTagWithOptions(field.key, { format, locale, fallback }));
      state.view = { kind: "main" };
    });
  });
  root.querySelector("#fo-cancel")!.addEventListener("click", () => {
    state.replaceTarget = null;
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- save + lint ----------

function renderSave(): HTMLElement {
  const revising = state.activeTemplate;
  const root = el(`
    <div class="section">
      <div class="section-head">Save to Salesforce</div>
      <div class="form">
        ${
          revising
            ? `<div class="banner info">Saving a new version of <b>${esc(revising.name)}</b>.
               <a href="#" id="sv-as-new">Save as a new template instead</a></div>`
            : ""
        }
        <label>Template name
          <input type="text" id="sv-name" value="${esc(revising?.name ?? suggestedTemplateName())}" />
        </label>
        <label>Base object
          <input type="text" value="${esc(state.baseObject ?? "")}" disabled />
        </label>
        <label>Test record Id (optional — used by Preview)
          <input type="text" id="sv-test" placeholder="001…" value="${esc(state.lastTestRecordId)}" />
        </label>
        <div class="btn-row">
          <button class="btn primary" id="sv-save">${revising ? "Save new version" : "Save template"}</button>
          <button class="btn secondary" id="sv-cancel">Cancel</button>
        </div>
      </div>
    </div>`);
  root.querySelector("#sv-as-new")?.addEventListener("click", (e) => {
    e.preventDefault();
    state.activeTemplate = null;
    state.view = { kind: "save" };
    render();
  });
  root.querySelector("#sv-save")!.addEventListener("click", () => {
    const name = root.querySelector<HTMLInputElement>("#sv-name")!.value.trim();
    if (!name) {
      state.error = "Template name is required.";
      render();
      return;
    }
    const testId = root.querySelector<HTMLInputElement>("#sv-test")!.value;
    void onSave(name, testId);
  });
  root.querySelector("#sv-cancel")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  return root;
}

function renderLintResults(result: SaveTemplateResponse): HTMLElement {
  const statusBanner =
    result.validationStatus === "Valid"
      ? `<div class="banner ok">Template saved and valid — ready to publish from the Sliick Docs Template Library.</div>`
      : `<div class="banner error">Template saved with problems — fix the tags below, then save again.</div>`;
  const items = result.tagCatalog
    .map((t) => {
      const suggest = t.suggestion
        ? `<div class="lint-suggest">Did you mean <b>${esc(t.suggestion)}</b>?</div>`
        : "";
      return `<li><div><span class="lint-tag">${esc(t.tag)}</span>${suggest}</div>
        <span class="lint-status ${t.status}">${t.status}</span></li>`;
    })
    .join("");
  const warnings = result.warnings
    .map((w) => `<div class="banner info">${esc(w.message)}</div>`)
    .join("");
  // PDF readiness (office-pdf). Only shown when the org supports PDF output and
  // the backend returned a verdict. Word output is unaffected either way.
  const pdfSupported = state.capabilities?.features.pdfOutput === true;
  let pdfBanner = "";
  if (pdfSupported && result.pdfReady !== undefined) {
    if (result.pdfReady) {
      pdfBanner = `<div class="banner ok">PDF-ready — this template can also generate PDF.</div>`;
    } else {
      const reasons = (result.pdfWarnings ?? [])
        .map((w) => `<li>${esc(w)}</li>`)
        .join("");
      pdfBanner =
        `<div class="banner info"><b>Word only for now.</b> This template has features that won't render in native PDF` +
        (reasons ? `:<ul class="lint-list">${reasons}</ul>` : ".") +
        ` It still generates Word perfectly.</div>`;
    }
  }
  const root = el(`
    <div class="section">
      <div class="section-head">Save results</div>
      <div class="form">
        ${statusBanner}
        ${pdfBanner}
        ${warnings}
        ${result.tagCatalog.length > 0 ? `<ul class="lint-list">${items}</ul>` : ""}
        <label>Preview record Id (optional — defaults to the template's test record)
          <input type="text" id="lr-record" placeholder="001…" value="${esc(state.lastTestRecordId)}" />
        </label>
        <div class="btn-row">
          <button class="btn secondary" id="lr-back">Back to fields</button>
          <span class="spacer"></span>
          <button class="btn primary" id="lr-preview" ${state.settings.mockMode ? "disabled title='Preview needs a Salesforce connection'" : ""}>Preview</button>
        </div>
      </div>
    </div>`);
  root.querySelector("#lr-back")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  root.querySelector("#lr-preview")!.addEventListener("click", () => {
    const recordId = root.querySelector<HTMLInputElement>("#lr-record")!.value;
    void onPreview(recordId);
  });
  return root;
}
