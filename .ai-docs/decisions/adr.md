# ADRs — Sliick Docs Office Add-in (client side)

Status: Accepted, shipped on `main` (2026-06-12).
Backend counterpart: `sliick-docs/.ai-docs/decisions/office-templates-adr.md`.
Contract of record: `sliick-docs/.ai-docs/plan/sprint-office-addin-backend/
functional-requirements.md` §4 — `src/api/types.ts` mirrors it 1:1; change the
doc first, then both sides.

## ADR-1: Office.js web add-in, no framework

One codebase for Word (Excel/PPT later) across Windows/Mac/web. Task pane is
vanilla TypeScript + Vite — single-state-object + full re-render; a framework
would cost more than it buys at this size. VSTO/COM rejected (Windows-only).

## ADR-2: Tags are plain text; the add-in is convenience, not a gatekeeper

The engine parses `{{...}}` out of OOXML regardless of how it got there.
Hand-typed tags work (degraded PDF-Butler-style mode); the add-in adds the
picker, wizards, and lint. Tags are inserted in a SINGLE `insertText` call so
each lands in one `<w:r>` run — the engine's run-coalescer is the fallback for
hand edits, not the primary path.

## ADR-3: Grammar emitted by the wizards matches the engine exactly

- `{{:else}}` (not `{{else}}`); single-quoted string literals in `{{#if}}`
  (the engine lexer rejects double quotes; embedded single quotes stripped —
  no escape syntax exists). Caught by cross-repo contract testing.
- Loop-table wizard emits Phase H row scope: `{{#Rel}}` opens in the first
  cell, `{{/Rel}}` closes in the last cell of the same row; in-loop keys are
  child-relative (`{{FirstName}}`).
- Wizards render only for features `GET /office/v1/capabilities` advertises —
  older subscriber orgs degrade gracefully.

## ADR-4: Auth = ECA + OAuth PKCE in an Office dialog

Public client (no secret in a browser app), S256 PKCE (RFC 7636 — verifier
unit-tested against the appendix B vector), login via
`Office.context.ui.displayDialogAsync` → `auth-callback.html` relays
code/state via `messageParent`. Access token in sessionStorage, refresh token
in localStorage; API client does one 401 → refresh → retry.

## ADR-5: Two-step save

`getFileAsync` (compressed, sliced) → base64 → standard ContentVersion REST
(step 1) → `POST /office/v1/templates` attaches + validates (step 2). Reason:
Apex REST's 6 MB sync cap vs the 10 MB file limit base64-inflated. The lint
panel renders the returned tag catalog verbatim.

## ADR-6: Mock mode implements the same contract interface

`MockSliickClient` implements `SliickApi` with scope-aware tag classification
(in-loop keys resolve against the child object — matching backend semantics),
so the full UX is demoable in Word with zero Salesforce. Demo mode is the
default until an org is configured.

## ADR-7: Editor-parity sprint decisions (2026-07)

- **Nested loops insert as paragraph-scope blocks, not nested tables** — each
  `{{#Rel}}`/`{{/Rel}}` on its own paragraph. Word.js nested-table insertion is
  brittle, and paragraph scope is a first-class engine shape (OfficeLoopScope).
- **Approvals is special-cased client-side**: the loop wizard skips
  `discover('ProcessInstanceStep')` and offers the fixed seven-field set
  mirroring `TemplateMergeFieldService.APPROVALS_FIELDS` — a generic discover
  would offer raw fields the engine rejects. No totals over Approvals (the
  fetch plan skips synthetic relationships).
- **No truthy-block builder** — the engine's `{{#X}}` is strictly a loop-open;
  `truthyTags` was removed so nobody wires it to UI. "Show when present" is
  `{{#if Field != null}}`.
- **Saves version the active template**: after any save (or picking one in
  My templates), subsequent saves send `templateId` and create new versions
  instead of name-collision duplicates. "Save as new" clears it; switching
  base object clears it.
- **Negated clauses are parenthesized** (`NOT (A = 'B')`) so Pratt-parser
  precedence can't misread them; `contains` literals are always quoted.

## ADR-8: Grammar-v2 wizards + UX overhaul decisions (2026-07)

- **One tag parser (`tagParse.ts`), inverse of the builders** — feeds local
  lint, the tag navigator, edit-in-wizard prefill, and cursor-scope
  detection. Mock lint delegates to it, so demo mode can't drift from the
  grammar. Round-trip tests pin builder↔parser symmetry.
- **Local lint is a courtesy, backend is the authority** (ADR-2 holds).
  `lint.ts` classifies scope-aware against cached discover data; loops whose
  child discover isn't fetched lint leniently. Structural OOXML checks
  (in-cell nested loops, cross-paragraph splits) stay server-side only.
- **Tag highlighting is authoring-only** — Word highlight applied per-tag via
  search, and ALWAYS stripped in onSave before getFileAsync, because run
  formatting survives the merge into generated output. No engine change.
- **Edit-in-wizard keeps plain text as the source of truth**: the selection
  probe hit-tests the cursor against parsed tags; wizards open prefilled and
  replace via select-then-insert (`replaceTarget`). Editable v1: scalar,
  barcode, loop-open (raw WHERE text + sort dropdowns). {{#if}} expressions
  are not parsed back into clauses (round-trip fidelity beats a lossy guess).
- **Selection probe is debounced (350 ms), main-view-only, best-effort** —
  it must never re-render a wizard under the user's typing or surface errors.
  Text-before-cursor comes from `body start → selection start` expandTo.
  **Size cutoff**: the first probe that reads >400k chars disables probing for
  the session (graceful absence of smart-cursor features on huge documents;
  the length is only knowable after a read, so we pay one big read then stop).
  If real-world docs need the features back at scale, the upgrade path is a
  paragraph-window read for tag hit-testing (tags can't span paragraphs) plus
  a short-TTL cache for the loop-scope prefix scan.
- **Nested wizard emits the Titan shape** (outer paragraphs + real inner
  table) — matches the engine regression added in sliick-docs grammar-v2.
- Recents/pins are per-base-object localStorage key lists; scroll position is
  captured/restored across full re-renders (same family as the wizard
  checkbox-state fix).

## Gotchas (don't re-learn)

- LWS blocks `URL.createObjectURL(Blob)` in Salesforce contexts; for downloads
  use `data:` URL anchors (the editor card does; the task pane runs outside
  LWS but the pattern is kept consistent).
- Office add-ins require HTTPS even on localhost — `office-addin-dev-certs`
  + Vite https config; manifest `<Version>` must be four-part ≥ 1.0.0.0.
- Icons must live in `public/` or they're served in dev but missing from the
  production build.
