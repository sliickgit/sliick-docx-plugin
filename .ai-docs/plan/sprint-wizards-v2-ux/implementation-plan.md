# Sprint: grammar-v2 wizards + UX overhaul (add-in)

Scope agreed 2026-07-02: wizard support for the grammar-v2 backend features
(loop filters/sort, fallback text, locale formats, barcodes, nested-as-table),
PLUS all four UX bundles: instant lint + tag navigator, edit-tag-in-wizard,
scope-aware fields + tag highlighting, starter layouts + polish batch.

Backend prerequisites are merged to sliick-docs main (5dca571): capabilities
`loopFilters` / `fallbackText` / `localeFormats` / `barcodes` all true.

## Design decisions

- **Tag parser is the keystone** (`src/office/tagParse.ts`): the inverse of the
  tag builders — one grammar-aware parser feeding local lint, the tag
  navigator, edit-in-wizard prefill, and cursor-scope detection. Pure TS,
  fully unit-tested against every builder output (round-trip property).
- **Local lint runs the same in demo and connected mode** (`src/office/lint.ts`):
  scope-aware classification against the discover data already in memory
  (child discovers fetched lazily + cached). It is a *courtesy* preview — the
  backend validator on save remains the authority (ADR-2).
- **Tag highlighting is authoring-only and auto-stripped**: a pane toggle
  applies Word highlight to `{{*}}` matches; `onSave` clears ALL tag
  highlights before `getFileAsync` so shading can never leak into generated
  documents. No engine change needed.
- **Edit-in-wizard is selection-driven**: DocumentSelectionChanged (debounced)
  reads the selection; when it sits inside/on a recognizable tag, the pane
  shows an "Edit this tag" card that opens the matching wizard prefilled and
  replaces the tag on confirm. Plain text stays the source of truth.
- **Cursor scope detection**: text before the cursor via
  `body.getRange('Start').expandTo(selection.getRange('Start'))`, scan for
  unclosed `{{#Rel}}` — pane banner + field list switch to the child object's
  in-loop fields.
- **Nested wizard emits the Titan shape**: outer loop paragraphs wrapping a
  real table (header + row-scope inner loop row) — engine regression
  `blockLoopWithInnerTableExpandsPerParent` covers it.

## Steps (sequential; verify = `npm test` + `npm run typecheck` green each step)

1. **Foundations** — types.ts capability flags; tags.ts: loop-open with
   WHERE/ORDER BY, `barcodeTag`, scalar with locale + fallback; unit tests.
2. **Tag parser** — tagParse.ts covering every grammar form incl. quote-aware
   WHERE/ORDER BY split; round-trip tests against the builders.
3. **Local lint** — lint.ts scope-aware classification (root/loop scopes,
   suggestions); mock.ts learns the new grammar (pipe, `*`, loop modifiers)
   + new capability flags; tests for both.
4. **Wizard sprint** — loop wizard filter/sort section; barcode wizard;
   field "insert with options" popover (format/fallback/locale); nested
   insert becomes paragraphs + real table (`insertNestedLoopWithTable`).
5. **Tags panel** — lint results list with per-tag status/suggestions,
   click-to-navigate (nth-occurrence search + select), refresh button;
   highlight-tags toggle + auto-strip in onSave.
6. **Selection smarts** — selection-changed handler (debounced): edit-tag
   card (prefills loop/cond/image/barcode/options wizards), in-loop scope
   banner + child-field list switch.
7. **Starter layout** — "Insert sample layout" building a demo invoice shape
   from live discover data (title, parent fields, Titan-shape nested table
   or flat table + total, conditional example).
8. **Polish batch** — recent/pinned fields (localStorage), Enter inserts top
   search hit, pane scroll preserved across re-renders, friendlier auth/API
   errors, Fluent-leaning style pass.
9. **Docs + ADR + full verify** — user guide updates, ADR-8 (parser/lint/
   highlight/selection decisions), `npm test` + `npm run build`; manual
   Word sideload checklist for Jerry (includes barcode visual check).

## Out of scope

- Any sliick-docs change (backend is done and merged).
- Excel/PowerPoint, template download/open-in-Word (no endpoint), Layout B.
