# Sprint: Editor parity (Tier 1) — implementation plan

Goal: close every gap between the add-in and the sliick-docs editor feature set that is
closable **in this repo alone** — i.e. the backend (package 1.8.0, grammar v1) already
supports the feature and the add-in simply lacks UI for it. Tier 2 gaps that need grammar
or engine work in sliick-docs are listed at the end as follow-ons, not built here.

Verified against the shipped engine before planning (2026-07-02):

- `OfficeAddinRestDiscover` returns the synthetic `Approvals` relationship
  (`childObjectApiName: 'ProcessInstanceStep'`), but the engine only accepts the fixed
  field set `ActorName, ActorTitle, StepStatus, Comments, ActedAt, ProcessName, StepName`
  (`TemplateMergeFieldService.APPROVALS_FIELDS`). A generic child `discover()` call would
  offer wrong columns → the loop wizard must special-case `Approvals`.
- `DocxTagEvaluator` has **no truthy-block kind** — `{{#X}}` is strictly `LOOP_OPEN`.
  The add-in's exported-but-unwired `truthyTags` builder would emit tags the validator
  rejects ("unknown relationship"). It gets **removed**, not wired up.
- Live capabilities (1.8.0): `nestedLoops:true`, `aggregates:true`, `picklistLabels:true`
  — the frozen contract doc §4 still shows the original `false` values and needs a
  doc-only correction in sliick-docs.
- No template-download endpoint exists in office/v1 → "open existing template in Word"
  is out of scope (contract-extension follow-on). The library view supports browse +
  save-new-version, which `SaveTemplateRequest.templateId` already enables.

## Steps (sequential, one session)

Each step lands with its unit tests; verify = `npm test` + `npm run typecheck` green
after every step, plus the step-specific check listed.

1. **Template library view + revise flow**
   New `templates` view listing `listTemplates(baseObject)` (client method already exists,
   currently unconsumed): name, validation badge, last-modified. Row action "Save new
   version" records `state.activeTemplate = {templateId, name}` and opens the save view
   with the name prefilled and a "saving a new version" notice + "save as new instead"
   escape hatch. `onSave` sends `templateId` when an active template is set. Mock
   `saveTemplate` honors `req.templateId` (update by id, keep id, new version id).
   *Verify:* mock unit tests — revise-by-id updates in place; list filters by base object;
   save-as-new leaves the original untouched.

2. **Preview against an ad-hoc record**
   Record-Id input on the save-results (lint) view, prefilled with the save-time
   `testRecordId`; `onPreview` sends `recordId` when present (contract field exists,
   currently never sent).
   *Verify:* typecheck + manual code path review (mock rejects preview by design).

3. **Conditional wizard: `contains` + NOT**
   The expression engine (`TemplateExpressionService`, grammar locked V1) supports
   `contains`, `NOT`, and parentheses; the wizard exposes neither. Add `contains` to both
   operator selects and a per-clause "NOT" checkbox. Builders emit `NOT (clause)` —
   parenthesized so parsing is unambiguous. `contains` values are always quoted.
   *Verify:* tags unit tests for `contains`, negated single clause, negated clause inside
   a compound expression; confirm emitted strings against the engine grammar (single
   quotes, exact keyword casing accepted by the lexer — check lexer for case handling).

4. **Nested loop wizard (depth-1)**
   Engine + capability flag support one nesting level; the wizard only builds flat tables.
   Extend the loop wizard: optional "Nested list" section (gated on
   `features.nestedLoops`) offering the child object's own child relationships (already
   fetched via the second `discover()` call — keep `childRelationships` from it) and a
   column picker. When chosen, insert a **paragraph-scope block** (not a nested table —
   Word.js nested-table insertion is brittle and the engine's paragraph scope is the
   documented nested shape):
   `{{#Rel}}` ¶ child fields line ¶ `{{#GrandRel}}` ¶ grandchild fields line ¶
   `{{/GrandRel}}` ¶ `{{/Rel}}`.
   New pure builder `nestedLoopBlockLines()` in tags.ts + `insertParagraphBlock()` in
   wordInsert.ts.
   *Verify:* tags unit tests — block line sequence, in-loop key stripping for both levels,
   single open/close per level, depth exactly 2.

5. **Approvals loop support**
   Loop wizard special-cases `relationshipName === 'Approvals'`: skip the child
   `discover()` call and use the fixed seven-field list (keys relative to loop scope,
   labels "Actor Name" etc.). Add the synthetic relationship + fields to the mock so demo
   mode exercises it. No aggregate section for Approvals (non-numeric fields).
   *Verify:* mock lint test — `{{#Approvals}}{{ActorName}}{{/Approvals}}` classifies
   Resolved; wizard uses fixed list (unit test on the helper).

6. **Housekeeping**
   Remove `truthyTags` (+ its tests) and the unused `insertImage` wrapper; fix the stale
   mock header comment ("no nested loops / aggregates" — both are true now). In
   sliick-docs: correct the §4 capabilities table in
   `.ai-docs/plan/sprint-office-addin-backend/functional-requirements.md` to the shipped
   1.8.0 values (doc-only; contract of record must match reality).
   *Verify:* grep shows no remaining references; full suite green.

7. **Docs + final verify**
   Update `docs/user-guide.md` for the new features. Full `npm test`,
   `npm run typecheck`, `npm run build`.
   Manual Word-side smoke (sideload) is deferred to Jerry — listed in the wrap-up.

## Out of scope (Tier 2 — need sliick-docs grammar/engine work first)

- Loop filters / sort (`{{#Rel WHERE …}}`), per-field fallback text, barcodes `{{*}}`,
  signature tags `{{@}}` merge behavior, rich-text field rendering, hyperlink fields,
  locale-override formats, header/footer image merge. Each needs a grammar extension +
  capability flag on the backend before add-in UI makes sense.
- Template download / open-in-Word (needs a new office/v1 endpoint).
- Excel / PowerPoint hosts (separate deferred track).
