# Sliick Docx Plugin — MS Office Merge-Field Plugin

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
| Format       | `{{Amount:currency}}`, `{{CloseDate:MM/dd/yyyy}}`, `{{Active:checkbox}}` |
| Built-ins    | `{{Today}}`, `{{Now}}`, `{{RunningUser.Name}}` |
| Loop (child) | `{{#Contacts}} … {{FirstName}} … {{/Contacts}}` |
| Conditional  | `{{#if Amount > 50000}} … {{else}} … {{/if}}`; truthy `{{#IsActive}}…{{/IsActive}}` |
| Aggregate    | `{{SUM:OpportunityLineItems.TotalPrice}}`, `{{COUNT:Contacts}}` |
| Image        | `{{%Logo__c:200x60}}` |

Phase 2 implements scalar/format/built-in/loop. Conditionals, aggregates, images
follow within Phase 2/4.

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
- **`OfficeMergeEngine` (Apex):** run-normalization (merge adjacent `<w:r>` so Word
  can't split a tag across runs), then scalar tags + formats + built-ins +
  `RunningUser`, resolving values via existing
  `TemplateRenderHelper.resolveMergeFieldRawValue()`.
  *Verify:* Apex test merges a doc with N scalar tags against a test record; output
  XML contains the resolved values.
- **Loops** over child relationships (both single-table-row and block forms); build
  child SOQL from discovery metadata. *Verify:* parent + 3 children → 3 rows.
- **Conditionals, aggregates, images.** *Verify:* one Apex test per feature.
- Apply the merge to **headers/footers** parts too, not just `document.xml`.

### Phase 3 — Upload, lifecycle & generation wiring
- **Data model:** add `Template_Format__c` picklist to `Template__c`
  (`NATIVE | OFFICE_WORD | OFFICE_EXCEL | OFFICE_POWERPOINT`); for Office formats,
  store the uploaded binary as a **ContentVersion** referenced by
  `Template_Version__c`; the `Template_Release__c` snapshot pins the binary for
  immutability. *(Use `sf-metadata`.)*
- **`POST /services/apexrest/sliick/office/v1/upload`** — create/update the Office
  template + store binary. Add-in "Save to Salesforce" button calls it.
  *Verify:* upload creates `Template__c` + `Template_Version__c` + binary; appears in
  the Template Library.
- **Single-record generation:** `OfficeRuntimeService.generate(templateId, recordId)`
  (parallel to `TemplateRuntimeService`), output saved via `FilesArtifactStore`,
  history via `Job__c`. *Verify:* Generate on an Account → merged `.docx` in Files.
- **Bulk:** branch `DocBatchRunner` on `Template_Format__c` to invoke the Office
  engine; reuse the whole batch lifecycle. *Verify:* a small bulk run produces N
  merged files.

### Phase 4 — Hardening + Excel/PowerPoint
- Malformed-tag error surfacing (snippet + location); FLS/sharing/`USER_MODE` review
  *(use `sf-security`)*; heap/governor review for large docs (note portwood's
  giant-query pattern as prior art if loops get huge).
- Optional robustness: insert tags as **Word content controls** (tag stored in the
  control) to eliminate run-splitting entirely — evaluate vs. plain text.
- **Excel** (`xl/worksheets/sheetN.xml` + `sharedStrings.xml`) and **PowerPoint**
  (`ppt/slides/slideN.xml`) reuse the plugin shell, auth, upload path, and tag
  grammar; only OOXML traversal differs. *Verify:* parallel per-format Apex tests +
  a round-trip generation each.

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
- **PDF output of Office templates:** Salesforce cannot natively render `.docx → .pdf`
  (`Blob.toPdf` takes HTML). Office templates therefore output their **native format**
  only. If PDF output of Office templates is required, it needs an external conversion
  service — out of scope for now; confirm this is acceptable.
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
