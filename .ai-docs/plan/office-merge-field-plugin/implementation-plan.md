# Sliick Docx Plugin — MS Office Merge-Field Plugin

> **Alignment update (2026-06-12).** sliick-docs already has a planned engine
> sprint — `sliick-docs/.ai-docs/plan/sprint-docx-input/functional-requirements.md`
> (Phase H: uploaded `.docx` templates, `{{...}}` grammar, run coalescing, tag
> evaluator, upload validation, `Source_Type__c='Uploaded'`) — which explicitly
> deferred the in-Word add-in. **This repo is that add-in.** The backend work for
> this plugin is now planned separately as
> `sliick-docs/.ai-docs/plan/sprint-office-addin-backend/functional-requirements.md`
> (ECA + CORS + REST contract wrapping Phase H services) and will be built on a
> sliick-docs feature branch independently. Supersessions for this doc:
> - Backend class names: Phase H's `TemplateDocxMergeService` / `DocxRunCoalescer`
>   / `DocxTagEvaluator` / `TemplateUploadedDocxValidator` replace the
>   `OfficeMergeEngine` / `OfficePdfSafetyLint` placeholders below.
> - Else token is **`{{:else}}`** (not `{{else}}`).
> - Engine v1 constraints the add-in must respect (gate wizards via the
>   `capabilities` endpoint): no nested loops, no aggregates, max 1 parent hop
>   inside a repeat, loops are paragraph- or single-table-row-scope, 10 MB cap.
>   Aggregates/nested loops/barcodes/signature tags are capability-flagged
>   follow-ons (backend plan §8.2).
> - Upload is **two-step** (standard ContentVersion REST for the binary, then
>   `POST /office/v1/templates` to attach + validate) because Apex REST caps
>   sync payloads at 6 MB.
> - PDF output of uploaded templates is a follow-on backend sprint (§8.1 there);
>   the PDF strategy section below remains the design of record for it.

## Context

`sliick-docs` is a Salesforce managed package (namespace `sliick`, API v66) that
generates documents from Salesforce records. Today templates are authored **only**
in a web WYSIWYG editor and stored as JSON (`Template_Version__c.Editor_State__c`),
with merge fields written `{{Object.Field}}`. It renders to PDF (`Blob.toPdf`) and
to Word by assembling OOXML from scratch. It already has: dynamic field discovery
(`TemplateMergeFieldService.discover()`), bulk generation (batch console,
`DocBatchRunner`, `Job__c`), Files storage (`FilesArtifactStore`), and a REST
portability surface.

We want a **PDF-Butler-style** experience: users author templates **inside Word,
Excel, and PowerPoint**, drop in merge fields from a Salesforce-aware picker, then
upload the file to Salesforce where `sliick-docs` uses it as a template for one-off
and bulk generation. This preserves the user's exact Office formatting — the whole
reason to author in Office rather than the web editor.

`portwood-docgen` (a separate, open-source Salesforce package) already does Office
binary merging with single-brace `{...}` tags. **We use it only as a learning
reference — no code is copied.**

### Decisions locked with the user
- **Backend = process Office binaries directly** (not convert to sliick-docs JSON).
  The JSON model has no representation for spreadsheet cells or slides, so Excel/PPT
  require direct OOXML processing anyway; one engine serves all three formats. We
  reuse sliick-docs' render-agnostic orchestration (bulk, jobs, Files, discovery).
- **Plugin = Office.js web add-in.** One codebase for Word/Excel/PowerPoint across
  Windows, Mac, and web. (Not VSTO/COM — Windows-desktop-only.)
- **Scope = Word first**, then extend the same architecture to Excel & PowerPoint.
- **Merge syntax = `{{Object.Field}}`** to match the existing sliick-docs editor.
- **Auth = External Client App (ECA) + OAuth 2.0 PKCE.** (Connected Apps are
  deprecated.) Public client, no secret, PKCE.
- **PDF output = Salesforce-native `Blob.toPdf()` first (zero cost), external
  renderer later.** `Blob.toPdf()` only accepts HTML, so the native path is
  merged `.docx` → OOXML→HTML translation (new Apex) → `Blob.toPdf()`. Built
  behind a `PdfRenderer` abstraction so a hosted service (e.g. GCP +
  LibreOffice/Gotenberg) can later be swapped in for full-fidelity rendering
  without touching templates or the merge engine. Native PDF is Word-only;
  Excel/PPT output native format until the external renderer exists.

### Success criterion (end-to-end)
A user opens Word, signs into their Salesforce org via the task pane, picks
`Account` and inserts `{{Account.Name}}` and a `{{#Contacts}}…{{/Contacts}}` loop
table, clicks "Save to Salesforce," then from an Account record clicks Generate and
receives a `.docx` in Files with the data merged and the original formatting intact —
and the same template runs in a bulk batch over many Accounts.

---

## Architecture

Two codebases:

1. **`sliick-docx-plugin`** (this repo, currently empty) — the Office.js add-in
   (TypeScript). Built static bundle + add-in manifest hosted on **`sliick-astro`**
   (or any static host). Talks to Salesforce over OAuth + new REST endpoints.

2. **`sliick-docs`** (existing package) — new Apex: REST endpoints, the OOXML merge
   engine, data-model fields for Office templates, ECA metadata, and wiring into the
   existing generation/bulk paths.

```
Word/Excel/PPT  ──Office.js task pane──┐
   (insert {{tags}}, save file)        │  OAuth PKCE (ECA)
                                       ▼
                         sliick REST  /office/v1/discover   (field picker data)
                                      /office/v1/upload     (store template binary)
                                       │
sliick-docs ───────────────────────────┘
   Template__c (+ Template_Format__c)  ── binary stored as ContentVersion
   OfficeMergeEngine (new) ── ZipReader → normalize runs → resolve {{tags}} → ZipWriter
   reuses: TemplateRenderHelper.resolveMergeFieldRawValue / TemplateMergeFieldService.discover
   wired into: OfficeRuntimeService (single) + DocBatchRunner (bulk) + FilesArtifactStore + Job__c
```

---

## Merge-tag grammar (defined fresh on the `{{ }}` base)

Mirrors portwood's *capabilities* (good UX) with our syntax and our own
implementation. Field keys are 1:1 with `TemplateMergeFieldService.discover()`.

| Purpose      | Syntax |
|--------------|--------|
| Scalar       | `{{Account.Name}}`, parent hops `{{Account.Owner.Email}}` |
| Format       | `{{Amount:currency}}`, `{{CloseDate:MM/dd/yyyy}}`, `{{Active:checkbox}}`, `{{Stage:label}}` |
| Built-ins    | `{{Today}}`, `{{Now}}`, `{{RunningUser.Name}}` |
| Loop (child) | `{{#Contacts}} … {{FirstName}} … {{/Contacts}}` (nested OK; `{{#Approvals}}` synthetic) |
| Conditional  | `{{#if Amount > 50000}} … {{:else}} … {{/if}}`; truthy `{{#IsActive}}…{{/IsActive}}`; inverse `{{^HasDiscount}}…{{/HasDiscount}}` |
| Aggregate    | `{{SUM:OpportunityLineItems.TotalPrice}}`, `{{COUNT:Contacts}}` |
| Image        | `{{%Logo__c:200x60}}`, `{{%Image:0}}` (nth attached) |
| Barcode/QR   | `{{*ProductCode}}` (Code 128), `{{*Website:qr:200}}` — Tier 2 |
| Signature    | `{{@Signature:Role:Order:Type}}` — reserved; passed through unmerged until Sigil e-sign consumes it |

Phase 2 implements scalar/format/built-in/loop. Conditionals, aggregates, images
follow within Phase 2/4; barcode/QR and rich-text fields are fast-follow (see
Feature scope below).

---

## Learnings from portwood-docgen (adapted, not copied)

Portwood is reference-only (open source; **no code copied**). These are the
mechanisms and product lessons worth re-deriving in our own implementation, and
where our picker-first architecture lets us do better.

### Engine mechanics to re-derive
- **Run normalization is load-bearing.** Word splits `{{Account.Name}}` across
  multiple `<w:r>` runs (spell-check, rsids, formatting). Approach: extract all
  `<w:t>` texts → flat string + offset map → find tags whose braces land in
  different segments → merge those runs → rebuild XML back-to-front so earlier
  offsets stay valid. Same algorithm covers PPT (`<a:t>`) later. Known limit:
  tags split across *paragraphs* are not recoverable — lint for it instead.
  **Our advantage:** picker-inserted tags can be written as a single run (or a
  content control), so normalization is a fallback for hand-typed/edited tags,
  not the primary path.
- **One parser, one path.** Portwood's tag logic lives in 3 parallel resolution
  paths (sync, giant-query, parent-level) that must be hand-kept in sync — a
  documented source of regressions. We build ONE tokenizer/resolver used by
  every path (single, bulk, preview, lint) from day one.
- **Loop container expansion + sibling guard.** `{{#Child}}` inside a table row
  should repeat the `<w:tr>`; inside a numbered list, the `<w:p>`. Guard: if the
  candidate container holds tags from another loop, fall back to inline
  expansion (prevents two sibling loops eating each other's content).
  **Conditionals must NEVER trigger container expansion** (a false branch would
  emit nothing while the container's close tag remains → corrupt XML).
  **Our advantage:** the add-in's loop wizard marks the repeated region
  structurally (tags placed around the row by the wizard), so heuristics are a
  fallback for hand-authored templates, not the norm.
- **Parts enumeration:** merge `word/document.xml` + all `word/headerN.xml` /
  `word/footerN.xml`; pass through `styles.xml`/`numbering.xml` untouched;
  update `[Content_Types].xml` + `document.xml.rels` only when adding images.
- **Excel = two-pass:** inline `sharedStrings.xml` values into worksheet cells
  first, then merge, then drop sharedStrings. **PPT:** per-slide rels files;
  strip revision-info parts or PowerPoint shows a "repair" dialog.
- **Heap discipline:** check heap pressure every N loop iterations and fail over
  to an async/queueable path instead of dying mid-merge; for native-PDF images
  use *relative* `/sfc/servlet.shepherd/version/download/<id>` URLs (never
  blobs, never absolute URLs — Flying Saucer breaks on absolute). If interactive
  generation hits Apex heap walls on big templates, portwood's proven escape
  hatch is client-side ZIP assembly in the LWC (Apex merges XML strings, browser
  packs the ZIP) — keep as a contingency, not the v1 design.
- **Testing pattern:** expose a `@TestVisible` merge-XML-string entry point and
  test with small inline XML fixtures (fast, readable) instead of base64 ZIP
  fixtures; keep a handful of real `.docx` round-trip tests as integration
  checks.

### Product/UX lessons (and how we apply them)
- **Validate early AND late.** Portwood validates schema but not tag syntax at
  save — users hit malformed-tag errors only at generate time. We can beat this:
  the upload endpoint already has the OOXML open and org schema via
  `discover()`, so lint at upload for: unbalanced `{{#}}/{{/}}`, unknown
  fields, FLS-inaccessible fields, tags split across paragraphs, PDF-unsafe
  features. Keep portwood's *rich runtime errors* too (offending tag verbatim +
  snippet + location) — both, not either.
- **Word authoring traps to absorb into wizards + lint** (these cost portwood
  real support volume): AutoFit-to-Contents recalculates column widths on every
  save (`<w:tblGrid>` vs `<w:tcW>` disagreement = "phantom width" PDF bugs) —
  loop-table wizard should set **Fixed Column Width**; Track Changes must be
  resolved before upload (tracked markup renders as content) — lint detects
  `<w:ins>`/`<w:del>`; comments leave artifacts — lint; >10MB uploads (suggest
  image compression); custom fonts → native PDF falls back to Helvetica-class
  fonts — lint warns per the PDF-safe matrix.
- **Per-template pinned test record + preview.** sliick-docs already has a test
  record on `Template__c` — reuse it: "Preview with test record" button in the
  task pane renders the merged doc against the pinned record. Catch issues at
  authoring time, not first generation.
- **Document the output matrix honestly.** PDF vs DOCX asymmetries (fonts,
  hyperlinks, rich-text images) get a user-facing capability table; the lint
  references it. Portwood's experience: documented constraints + guidance cut
  support volume more than engine fixes did.
- **Bulk: estimate before submit, partial success always.** sliick-docs already
  has `DocBatchAnalyzerService` + per-record `Job__c` — extend both to Office
  templates rather than inventing anything.
- **Caution on output-format overrides:** portwood shipped then *removed* a
  runtime "Output As" override because cross-format output produced corrupt
  files. Our DOCX→PDF path is a deliberate translation layer (not a naive
  override), but the lesson stands: make PDF an explicit per-template
  capability gated by the PDF-safe lint, not a free toggle on every template.
- **Managed-package trap:** anything subscriber-facing (interfaces like a
  future `DataProvider`, Flow DTOs) must be `global` top-level classes with
  `@AuraEnabled` members and a `global` no-arg constructor — portwood shipped a
  feature unusable in subscriber orgs for a full release by missing this.

---

## Feature scope & parity matrix (vs portwood-docgen)

Parity strategy: portwood bundles engine + orchestration + e-sign in one package.
We split: **sliick-docs already owns** orchestration (bulk console + analyzer,
`Job__c` history, Files delivery, template library/folders/versions/releases,
watermark config, test record, bundles, Flow invocables, portability) and is
building e-signature natively (1.7.0 "Sigil"). This plugin + engine only needs the
**tag/render layer** for Office binaries. Don't rebuild what sliick-docs has.

**No user-authored query configs — ever.** Portwood's V1→V4 query-config evolution
(flat string → junction JSON → node tree → Apex provider) is four iterations of
making users describe their data twice. sliick-docs already solves this the right
way: derive the fetch plan from the merge-field catalog (`MergeFieldCatalog__c`).
Our upload lint extracts the catalog from the tags in the document; the engine
builds SOQL from it. The template IS the query spec. (A V4-style Apex
`DataProvider` interface remains a possible later escape hatch for computed/
external data — `global`, top-level, from day one if built.)

### Tier 1 — v1 (Word engine, in plan phases 2–3)
| Feature | Notes |
|---|---|
| Scalar tags + parent hops (≤5) | matches `discover()` keys |
| Format suffixes | `:currency`, `:percent`, `:number`, date patterns, `:checkbox`, `:label` (picklist label — cheap, high value) |
| Built-ins | `{{Today}}`, `{{Now}}`, `{{RunningUser.*}}` |
| Loops | nested, container expansion, sibling guard, empty-set safe |
| Conditionals | truthy, inverse `{{^}}`, `{{#if expr}}` with `=,!=,<,>,<=,>=`, `{{else}}` |
| Aggregates | `SUM/COUNT/AVG/MIN/MAX` outside loops |
| Images | `{{%Field}}` (ContentVersion id), sizing, `{{%Image:N}}` (nth attached) |
| Approvals loop | `{{#Approvals}}` — near-free: discovery already exposes the synthetic relationship |
| Headers/footers merge | + pass-through styles/numbering |
| Page breaks in loops / repeat header rows | native Word constructs — engine must not break them; PDF path maps repeat-header → `<thead>` |
| Dynamic file naming | merge tags in output title — adopt portwood's `Document_Title_Format` idea on `Template__c` if absent |
| Upload lint + pinned-test-record preview | our differentiator (portwood validates late) |

### Tier 2 — fast follow
| Feature | Notes |
|---|---|
| **QR codes + Code 128 barcodes** | pure-Apex PNG rasterization (own implementation). Improve on portwood: embed the PNG in **DOCX output too**, not PDF-only — we already have the image-embed path. Syntax `{{*Field:qr:200}}` / `{{*Field}}` |
| Rich-text-area fields | basic HTML→OOXML (bold/italic/lists); images via `{{%Field}}` |
| Hyperlink fields | clickable in DOCX output (portwood PDF renders text-only — match that limit natively) |
| Conditional operators `AND/OR/NOT` | extend `{{#if}}` grammar |
| Locale-aware formats | `:date:de_DE` style |

### Tier 3 — later / on demand
| Feature | Notes |
|---|---|
| Charts (`{{Chart:...}}`) | portwood's biggest engine investment (9 styles, pure-Apex PNG). Real differentiator but heavy; build only on customer demand |
| Async giant-loop path | contingency already noted (heap failover); build when a customer hits the wall |
| Apex DataProvider (V4-style) | escape hatch for external/computed data |
| Multi-language UI | follow sliick-docs' broader localization stance |
| Restricted-editing output regions | PDF-Butler-style; roadmap idea |

### Reserved now, delivered by sliick-docs later
- **Signature placement tag** — reserve `{{@Signature:Role:Order:Type}}` in the
  grammar v1 (parser recognizes + preserves it, lint validates it) so Office
  templates plug into Sigil e-signature when it ships. Engine must pass it
  through unmerged.

### Deliberately skipped (and why)
- HTML/Google-Docs templates, drag-and-drop builder → that's the existing
  sliick-docs web editor's job; the plugin exists for Office fidelity.
- Fillable-PDF AcroForm mapping → different product surface; revisit only with
  the external PDF renderer.
- User-facing query builders (V1–V4) → replaced by catalog-derived fetch plan.
- Combined-PDF merge modes for Office outputs → sliick-docs' existing
  bulk/merge handles PDFs it renders; DOCX outputs are individual files.

---

## Plan (step → verify)

### Phase 0 — Scaffolding & spikes
- **Add-in scaffold** in `sliick-docx-plugin`: Vite + TypeScript Office.js task-pane
  add-in, XML/unified manifest, dev sideload config.
  *Verify:* sideloaded in desktop **and** web Word; a button inserts a hardcoded
  `{{Account.Name}}` at the cursor via `Word.run`.
- **Apex OOXML spike** in `sliick-docs`: open a `.docx` with `Compression.ZipReader`,
  read `word/document.xml`, regex-replace one tag, repack with `Compression.ZipWriter`.
  *Verify:* an Apex test round-trips a real `.docx`; output opens in Word uncorrupted
  with the value merged. (Confirms `ZipReader` is available alongside the `ZipWriter`
  the package already uses.)

### Phase 1 — Auth + field discovery
- **ECA metadata** in `sliick-docs` (`ExternalClientApplication` +
  `ExtlClntAppOauthSettings`): scopes `api refresh_token openid`, PKCE on, public
  client, callback = add-in redirect page. *(Use `sf-metadata` skill.)*
- **OAuth PKCE flow** from the add-in via `Office.context.ui.displayDialogAsync`.
  *Verify:* add-in obtains a valid access token against a scratch/dev org and stores
  the refresh token.
- **`GET /services/apexrest/sliick/office/v1/discover`** wrapping
  `TemplateMergeFieldService.discover()` + child-relationship discovery.
  *(Use `sf-apex` + `sf-security`: `USER_MODE`, typed DTOs, FLS.)*
  *Verify:* picker tree renders real org objects → scalar fields, parent lookups,
  child relationships (for loops), and built-ins.

### Phase 2 — Word merge engine (core)
- **`OfficeMergeEngine` (Apex):** run-normalization first (flat-text + offset map;
  see Learnings), then a **single shared tokenizer/resolver** (used by every later
  path) for scalar tags + formats + built-ins + `RunningUser`, resolving values via
  existing `TemplateRenderHelper.resolveMergeFieldRawValue()`. Expose a
  `@TestVisible` merge-XML-string entry point; test with inline XML fixtures plus a
  few real `.docx` round-trips.
  *Verify:* Apex test merges a doc with N scalar tags against a test record —
  including a tag deliberately split across `<w:r>` runs; output XML contains the
  resolved values.
- **Loops** over child relationships: container expansion (`<w:tr>` / numbered
  `<w:p>`) with sibling-guard fallback to inline; conditionals never expand
  containers; build child SOQL from discovery metadata.
  *Verify:* parent + 3 children → 3 rows; two sibling loops in one table don't eat
  each other's content; empty child set emits nothing and leaves valid XML.
- **Conditionals, aggregates, images.** *Verify:* one Apex test per feature,
  including false-branch-leaves-valid-XML.
- Apply the merge to **headers/footers** (`headerN.xml`/`footerN.xml`) too; pass
  through `styles.xml`/`numbering.xml` untouched.

### Phase 3 — Upload, lifecycle & generation wiring
- **Data model:** add `Template_Format__c` picklist to `Template__c`
  (`NATIVE | OFFICE_WORD | OFFICE_EXCEL | OFFICE_POWERPOINT`); for Office formats,
  store the uploaded binary as a **ContentVersion** referenced by
  `Template_Version__c`; the `Template_Release__c` snapshot pins the binary for
  immutability. *(Use `sf-metadata`.)*
- **`POST /services/apexrest/sliick/office/v1/upload`** — create/update the Office
  template + store binary, running the **upload lint** (one pass, returns warnings):
  unbalanced `{{#}}/{{/}}`, unknown/FLS-inaccessible fields (via `discover()`),
  tags split across paragraphs, Track Changes / comment markup present, >10MB,
  plus the PDF-safe checks. Add-in "Save to Salesforce" shows warnings inline.
  *Verify:* upload creates `Template__c` + `Template_Version__c` + binary; a doc
  with an unclosed loop and an unknown field returns both warnings; appears in the
  Template Library.
- **Preview with pinned test record:** task-pane button merges the current doc
  against `Template__c`'s existing test record and returns the result for download.
  *Verify:* preview of a template with loops renders correct child rows.
- **Single-record generation:** `OfficeRuntimeService.generate(templateId, recordId)`
  (parallel to `TemplateRuntimeService`), output saved via `FilesArtifactStore`,
  history via `Job__c`. *Verify:* Generate on an Account → merged `.docx` in Files.
- **Bulk:** branch `DocBatchRunner` on `Template_Format__c` to invoke the Office
  engine; reuse the whole batch lifecycle. *Verify:* a small bulk run produces N
  merged files.
- **PDF output (native, Word-only):**
  - `PdfRenderer` interface + `NativePdfRenderer`: translate the **merged** OOXML to
    Flying-Saucer-safe HTML (CSS 2.1, tables-for-layout, `@page` regions for
    headers/footers, inline images only), then `Blob.toPdf()`. Reuse HTML/CSS
    conventions from the existing native render path where possible.
  - Output-format choice (`DOCX | PDF`) on the generation request; bulk honors it.
  - **PDF-safe lint at upload:** while the engine has the OOXML open, detect
    features that won't survive native PDF (text boxes, floating/absolutely
    positioned images, non-core fonts, SmartArt/shapes/charts) and return warnings
    to the add-in: renders fine as DOCX, degraded as native PDF, full-fidelity on
    the future external renderer. *Verify:* a template with a text box uploads with
    the expected warning; a "PDF-safe" template renders to PDF with layout intact.
  - *Verify:* Generate-as-PDF on an Account → PDF in Files with merged values; the
    same template generates DOCX unchanged.

### Phase 4 — Hardening + Excel/PowerPoint
- Malformed-tag error surfacing (snippet + location); FLS/sharing/`USER_MODE` review
  *(use `sf-security`)*; heap/governor review for large docs (note portwood's
  giant-query pattern as prior art if loops get huge).
- Optional robustness: insert tags as **Word content controls** (tag stored in the
  control) to eliminate run-splitting entirely — evaluate vs. plain text.
- **Excel** (`xl/worksheets/sheetN.xml` + `sharedStrings.xml`) and **PowerPoint**
  (`ppt/slides/slideN.xml`) reuse the plugin shell, auth, upload path, and tag
  grammar; only OOXML traversal differs. Excel: two-pass (inline shared strings
  into cells, then merge, then drop `sharedStrings.xml`). PPT: run normalization on
  `<a:t>`, per-slide `_rels` for images, strip revision-info parts (else PowerPoint
  shows a repair dialog). *Verify:* parallel per-format Apex tests + a round-trip
  generation each; merged `.pptx` opens with no repair prompt.

---

## Critical files

**New — `sliick-docx-plugin` (add-in):**
- `manifest.xml` — add-in manifest (host apps, task pane, permissions, hosted URL)
- `src/taskpane/` — task pane UI: auth, field-picker tree, insert actions
- `src/auth/` — OAuth PKCE dialog flow + token storage
- `src/office/` — `Word.run` / Excel / PPT insertion helpers
- `src/api/` — REST client for `discover` / `upload`

**New — `sliick-docs` (Apex + metadata):**
- `classes/OfficeMergeEngine.cls` — OOXML decompress → normalize → resolve → recompress
- `classes/OfficeRuntimeService.cls` — single-record generation entry point
- `classes/PdfRenderer.cls` (interface) + `classes/NativePdfRenderer.cls` — merged
  OOXML → CSS 2.1 HTML → `Blob.toPdf()`; future `ExternalPdfRenderer` (GCP callout)
- `classes/OfficePdfSafetyLint.cls` — upload-time detection of native-PDF-unsafe features
- `classes/OfficeTemplateRest.cls` — `/office/v1/discover` + `/office/v1/upload`
- `externalClientApps/…` — ECA + OAuth settings metadata
- `objects/Template__c/fields/Template_Format__c.field-meta.xml`

**Reused (do not reinvent):**
- `classes/TemplateRenderHelper.cls` — `resolveMergeFieldRawValue()` value resolution
- `classes/TemplateMergeFieldService.cls` — `discover()` field/relationship metadata
- `classes/DocBatchRunner.cls` — bulk orchestration (branch on format)
- `FilesArtifactStore` + `Job__c` — output storage + history

---

## Open questions / risks
- **Native PDF fidelity (`Blob.toPdf()` / Flying Saucer):** the native PDF path
  requires an OOXML→HTML translation and inherits CSS 2.1-era limits — no flex/grid,
  core fonts only, no floating/absolutely-positioned elements, headers/footers via
  `@page` regions, shapes/SmartArt dropped or rasterized. Mitigations: PDF-safe lint
  at upload (warn, don't fail), authoring guidance generated by the add-in wizards
  (tables not text boxes, inline images), and the `PdfRenderer` abstraction so a
  hosted GCP service (LibreOffice/Gotenberg-class) later provides full-fidelity
  `.docx → .pdf` with zero template changes. Native PDF is Word-only; Excel/PPT are
  native-format output until the external renderer ships.
- **Run-splitting in Word:** mitigated by engine-side run normalization in v1; content
  controls are the belt-and-suspenders option (Phase 4).
- **Distribution:** sideload/admin-deployed for internal testing; AppSource listing is
  a later, separate track.
- **Large documents / many child rows:** watch Apex heap and CPU; portwood's
  giant-query/queueable pattern is prior art if needed.

---

## Verification (end-to-end)
1. Sideload the add-in in desktop + web Word; sign in via ECA OAuth against a scratch
   org.
2. Pick `Account`; insert `{{Account.Name}}`, `{{Account.Owner.Email}}`, a
   `{{#Contacts}}…{{/Contacts}}` table, and `{{Today}}`.
3. Save to Salesforce; confirm `Template__c` (format `OFFICE_WORD`) + binary in the
   Template Library.
4. From an Account record, Generate → open the resulting `.docx` from Files: values
   merged, child rows expanded, original formatting intact.
5. Run a bulk batch over several Accounts → N correct files, `Job__c` history present.
6. Apex test suite green for `OfficeMergeEngine` (scalars, formats, loops,
   conditionals, aggregates, images) and the REST endpoints (with non-admin FLS path).
