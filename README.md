# Sliick Docs for Microsoft Office

Office.js task-pane add-in that lets users build **Sliick Docs templates inside
Microsoft Word**: browse a Salesforce org's merge-field catalog, insert
`{{Object.Field}}` tags, repeating tables, and conditionals, then save the
document to Salesforce as an Uploaded template — with inline tag validation.

Companion to the [sliick-docs](../sliick-docs) managed package. The Salesforce
side of this integration is specified in
`sliick-docs/.ai-docs/plan/sprint-office-addin-backend/functional-requirements.md`
(REST contract v1) and builds on the `sprint-docx-input` engine sprint.
The add-in plan lives at
[.ai-docs/plan/office-merge-field-plugin/implementation-plan.md](.ai-docs/plan/office-merge-field-plugin/implementation-plan.md).

## Status

- **Word**: full task pane — field picker, repeating-table wizard, conditional
  wizards, save + lint results, preview hook. Excel/PowerPoint: planned.
- **Demo mode** (default): runs against an in-memory mock of the REST contract —
  full UX in Word with sample Account/Contact/Opportunity data, including
  scope-aware tag linting on save. No Salesforce needed.
- **Connected mode**: OAuth 2.0 PKCE against an External Client App + the
  `/services/apexrest/sliick/office/v1/*` endpoints. Functional once the
  backend sprint ships in the org.

## Quick start (dev)

```bash
npm install
npm run certs      # one-time: install trusted localhost HTTPS certs (required by Office)
npm run dev        # serves https://localhost:3000
```

Then sideload into Word:

```bash
npm run sideload   # opens desktop Word with the add-in registered
```

(or manually: Word → Insert → Add-ins → upload `manifest.xml`; on Word for the
web use *Add-ins → Upload My Add-in*.)

Open the **Merge Fields** button on the Home ribbon. The pane starts in **Demo
mode** — pick *Account*, click fields to insert tags, use *Related lists* to
insert a repeating Contacts table, then **Save to Salesforce** to see the lint
panel work against the document's real text.

To connect a real org (once the backend is installed): ⚙ Settings → untick Demo
mode → enter the org's My Domain URL and the ECA consumer key → Sign in.

## Scripts

| Script | What |
|---|---|
| `npm run dev` | Vite dev server (https if certs installed) |
| `npm run build` | Typecheck + production build to `dist/` |
| `npm test` | Vitest unit suite (PKCE, tag grammar, mock lint) |
| `npm run validate` | Validate `manifest.xml` |
| `npm run sideload` / `sideload:stop` | Register/unregister in desktop Office |
| `npm run certs` | Install trusted localhost dev certificates |

## Architecture

```
taskpane.html / auth-callback.html      Vite entries
src/
  taskpane/    UI shell (vanilla TS, single-state render)
  office/      All Word.run document mutations + pure tag builders (tags.ts)
  api/         types.ts = frozen REST contract DTOs (§4 backend plan)
               client.ts = live org client (bearer + 401-refresh-retry,
                           two-step ContentVersion upload)
               mock.ts   = contract mock w/ scope-aware tag classification
  auth/        PKCE utils, ECA OAuth dialog flow, settings storage
```

Design rules:

- **Tags are plain text** inserted in a single run — the add-in is convenience
  over a format, never a gatekeeper. Hand-typed tags work too; the engine's
  run-coalescer and the save-time lint catch the mess Word makes.
- **Loop tables follow Phase H row scope**: `{{#Rel}}` opens in the first cell,
  `{{/Rel}}` closes in the last cell of the same row.
- **Capabilities gating**: wizards render only for features the installed
  package version reports (`GET /office/v1/capabilities`), so the add-in
  degrades gracefully in older subscriber orgs.
- The contract DTOs in `src/api/types.ts` change **only** together with the
  backend plan document — they are the interface between the two repos.

## Production hosting

`npm run build` emits a static `dist/` (two pages + assets). Host it on any
static host (planned: sliick-astro site), update the URLs in `manifest.xml`
from `https://localhost:3000` to the production origin, and distribute the
manifest via Microsoft 365 admin center (Integrated Apps) or AppSource.
