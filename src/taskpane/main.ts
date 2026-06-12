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
  loadSettings,
  normalizeOrgUrl,
  saveSettings,
} from "../auth/settings";
import { defaultFormatForType, inLoopFieldKey, scalarTag } from "../office/tags";
import {
  getDocumentAsBase64,
  getDocumentText,
  insertConditional,
  insertInverse,
  insertLoopTable,
  insertScalar,
  suggestedTemplateName,
} from "../office/wordInsert";

// ---------------------------------------------------------------- state

type View =
  | { kind: "main" }
  | { kind: "settings" }
  | { kind: "connect" }
  | { kind: "loopWizard"; rel: ChildRelationshipDef; childFields: MergeFieldDef[] }
  | { kind: "condWizard"; mode: "if" | "inverse" }
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
};

// ---------------------------------------------------------------- boot

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    document.getElementById("app")!.innerHTML =
      `<div class="pane"><div class="banner error">Sliick Docs currently supports Word. Excel and PowerPoint are coming next.</div></div>`;
    return;
  }
  void initApi();
});

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
        state.settings.clientId,
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
    state.error = e instanceof Error ? e.message : String(e);
  } finally {
    state.busy = null;
    render();
  }
}

async function onInsertField(field: MergeFieldDef): Promise<void> {
  const format = defaultFormatForType(field.type);
  await withBusy(null, async () => {
    await insertScalar(scalarTag(field.key, format));
  });
}

async function onLoopWizardOpen(rel: ChildRelationshipDef): Promise<void> {
  await withBusy(`Loading ${rel.label} fields…`, async () => {
    const childDiscover = await state.api!.discover(rel.childObjectApiName);
    state.view = {
      kind: "loopWizard",
      rel,
      childFields: childDiscover.rootScalarMergeFields,
    };
  });
}

async function onLoopInsert(rel: ChildRelationshipDef, picked: MergeFieldDef[]): Promise<void> {
  await withBusy("Inserting table…", async () => {
    await insertLoopTable(
      rel.relationshipName,
      picked.map((f) => ({ inLoopKey: inLoopFieldKey(f.key), label: f.label })),
    );
    state.view = { kind: "main" };
  });
}

async function onSave(name: string, testRecordId: string): Promise<void> {
  await withBusy("Saving to Salesforce…", async () => {
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
        ...(testRecordId.trim() ? { testRecordId: testRecordId.trim() } : {}),
      },
      fileBase64,
    );
    state.lastSavedVersionId = result.versionId;
    state.view = { kind: "lintResults", result };
  });
}

async function onPreview(): Promise<void> {
  if (!state.lastSavedVersionId) return;
  await withBusy("Generating preview…", async () => {
    const blob = await state.api!.preview({ versionId: state.lastSavedVersionId! });
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
    state.tokens = await login(state.settings.orgUrl, state.settings.clientId);
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

function esc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function render(): void {
  const app = document.getElementById("app")!;
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
      pane.appendChild(renderLoopWizard(state.view.rel, state.view.childFields));
      break;
    case "condWizard":
      pane.appendChild(renderCondWizard(state.view.mode));
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
      if (!clientInput.value.trim()) {
        state.error = "Consumer key is required to connect to Salesforce.";
        render();
        return;
      }
    }
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
  const configured = state.settings.orgUrl && state.settings.clientId;
  const root = el(`
    <div class="section">
      <div class="section-head">Connect to Salesforce</div>
      <div class="form">
        ${
          configured
            ? `<p class="hint">Sign in to <b>${esc(state.settings.orgUrl)}</b> to load your org's merge fields.</p>
               <button class="btn primary" id="btn-login">Sign in to Salesforce</button>`
            : `<p class="hint">Set your org URL and consumer key in Settings (⚙) first — or switch on Demo mode to explore with sample data.</p>`
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

  // Search
  const search = el(
    `<input class="search" type="text" placeholder="Search merge fields…" value="${esc(state.search)}" />`,
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
  pane.appendChild(search);

  const d = state.discover;
  if (!d) {
    if (!state.busy) pane.appendChild(el(`<div class="hint">Pick a base object to load fields.</div>`));
    return;
  }

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
  actions.appendChild(el(`<span class="spacer"></span>`));
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
  for (const f of visible) {
    const row = el(`
      <button class="field-row" title="Insert ${esc(scalarTag(f.key, defaultFormatForType(f.type)))}">
        <span class="f-label">${esc(f.label)}</span>
        <span class="f-type">${esc(f.type)}</span>
      </button>`);
    row.addEventListener("click", () => void onInsertField(f));
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

function renderLoopWizard(rel: ChildRelationshipDef, childFields: MergeFieldDef[]): HTMLElement {
  const checks = childFields
    .map(
      (f, i) =>
        `<label><input type="checkbox" data-i="${i}" ${i < 3 ? "checked" : ""}/> ${esc(f.label)} <span class="f-type">${esc(f.type)}</span></label>`,
    )
    .join("");
  const root = el(`
    <div class="section">
      <div class="section-head">Repeating table — ${esc(rel.label)}</div>
      <div class="form">
        <p class="hint">Pick the columns. A table is inserted where your cursor is:
        a header row plus one row that repeats for every ${esc(rel.label)} record.</p>
        <div class="col-list">${checks}</div>
        <div class="btn-row">
          <button class="btn primary" id="lw-insert">Insert table</button>
          <button class="btn secondary" id="lw-cancel">Cancel</button>
        </div>
      </div>
    </div>`);
  root.querySelector("#lw-insert")!.addEventListener("click", () => {
    const picked: MergeFieldDef[] = [];
    root.querySelectorAll<HTMLInputElement>("input[type=checkbox]").forEach((box) => {
      if (box.checked) {
        const idx = Number(box.dataset.i);
        const f = childFields[idx];
        if (f) picked.push(f);
      }
    });
    if (picked.length === 0) {
      state.error = "Pick at least one column.";
      render();
      return;
    }
    void onLoopInsert(rel, picked);
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
                 </select>
               </label>
               <label>Value <input type="text" id="cw-value" placeholder="e.g. 50000 or Closed Won" /></label>
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
    const op = root.querySelector<HTMLSelectElement>("#cw-op")!.value as
      | "=" | "!=" | ">" | "<" | ">=" | "<=";
    const value = root.querySelector<HTMLInputElement>("#cw-value")!.value.trim();
    if (!value) {
      state.error = "Enter a value to compare against.";
      render();
      return;
    }
    const fieldType = fieldSel.selectedOptions[0]?.dataset.type ?? "string";
    const numericTypes = new Set(["currency", "int", "double", "percent"]);
    void withBusy("Inserting…", async () => {
      await insertConditional({
        fieldKey,
        operator: op,
        value,
        quoteValue: !numericTypes.has(fieldType),
        withElse: root.querySelector<HTMLInputElement>("#cw-else")!.checked,
      });
      state.view = { kind: "main" };
    });
  });
  root.querySelector("#cw-cancel")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  return root;
}

// ---------- save + lint ----------

function renderSave(): HTMLElement {
  const root = el(`
    <div class="section">
      <div class="section-head">Save to Salesforce</div>
      <div class="form">
        <label>Template name
          <input type="text" id="sv-name" value="${esc(suggestedTemplateName())}" />
        </label>
        <label>Base object
          <input type="text" value="${esc(state.baseObject ?? "")}" disabled />
        </label>
        <label>Test record Id (optional — used by Preview)
          <input type="text" id="sv-test" placeholder="001…" />
        </label>
        <div class="btn-row">
          <button class="btn primary" id="sv-save">Save template</button>
          <button class="btn secondary" id="sv-cancel">Cancel</button>
        </div>
      </div>
    </div>`);
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
  const root = el(`
    <div class="section">
      <div class="section-head">Save results</div>
      <div class="form">
        ${statusBanner}
        ${warnings}
        ${result.tagCatalog.length > 0 ? `<ul class="lint-list">${items}</ul>` : ""}
        <div class="btn-row">
          <button class="btn secondary" id="lr-back">Back to fields</button>
          <span class="spacer"></span>
          <button class="btn primary" id="lr-preview" ${state.settings.mockMode ? "disabled title='Preview needs a Salesforce connection'" : ""}>Preview with test record</button>
        </div>
      </div>
    </div>`);
  root.querySelector("#lr-back")!.addEventListener("click", () => {
    state.view = { kind: "main" };
    render();
  });
  root.querySelector("#lr-preview")!.addEventListener("click", () => void onPreview());
  return root;
}
