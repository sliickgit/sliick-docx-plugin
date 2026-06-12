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

## Gotchas (don't re-learn)

- LWS blocks `URL.createObjectURL(Blob)` in Salesforce contexts; for downloads
  use `data:` URL anchors (the editor card does; the task pane runs outside
  LWS but the pattern is kept consistent).
- Office add-ins require HTTPS even on localhost — `office-addin-dev-certs`
  + Vite https config; manifest `<Version>` must be four-part ≥ 1.0.0.0.
- Icons must live in `public/` or they're served in dev but missing from the
  production build.
