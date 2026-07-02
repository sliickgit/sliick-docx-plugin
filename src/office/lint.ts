/**
 * Local tag lint — instant, in-pane checking of the document's tags against
 * the org's discover metadata, in both demo and connected mode. This is a
 * courtesy preview for the author (typos, wrong scope, unknown lists); the
 * backend validator at save time remains the authority (ADR-2) — notably the
 * structural row/cell checks (in-cell nested loops, cross-paragraph splits)
 * only run server-side where OOXML structure is visible.
 */

import { DiscoverResponse, MergeFieldDef } from "../api/types";
import { inLoopFieldKey } from "./tags";
import { findTags } from "./tagParse";

/**
 * The synthetic Approvals relationship only accepts this fixed field set —
 * mirrors TemplateMergeFieldService.APPROVALS_FIELDS in sliick-docs.
 */
export const APPROVALS_LOOP_FIELDS: MergeFieldDef[] = [
  { key: "ActorName", label: "Actor Name", type: "string" },
  { key: "ActorTitle", label: "Actor Title", type: "string" },
  { key: "StepStatus", label: "Step Status", type: "string" },
  { key: "Comments", label: "Comments", type: "string" },
  { key: "ActedAt", label: "Acted At", type: "datetime" },
  { key: "ProcessName", label: "Process Name", type: "string" },
  { key: "StepName", label: "Step Name", type: "string" },
];

interface LintScope {
  label: string;
  keys: Set<string>; // lowercase
  proper: Map<string, string>; // lowercase → proper-cased key
}

export interface LintContext {
  root: LintScope;
  /** lowercase relationship name → in-loop scope (bare child keys). */
  childScopes: Map<string, LintScope>;
  /** lowercase relationship name → proper-cased name. */
  relationships: Map<string, string>;
}

function scopeFromFields(label: string, fields: MergeFieldDef[]): LintScope {
  const keys = new Set<string>();
  const proper = new Map<string, string>();
  for (const f of fields) {
    keys.add(f.key.toLowerCase());
    proper.set(f.key.toLowerCase(), f.key);
  }
  return { label, keys, proper };
}

/**
 * Builds the lint context from the base object's discover plus any child
 * discovers already fetched (keyed by childObjectApiName). Relationships
 * whose child discover isn't loaded yet lint leniently (in-loop fields pass).
 */
export function buildLintContext(
  root: DiscoverResponse,
  childDiscovers: Map<string, DiscoverResponse>,
): LintContext {
  const rootScope = scopeFromFields(root.baseObjectLabel, [
    ...root.rootScalarMergeFields,
    ...root.parentLookupMergeFields,
    ...root.runningUserMergeFields,
    ...root.builtInMergeFields,
  ]);
  const childScopes = new Map<string, LintScope>();
  const relationships = new Map<string, string>();
  for (const rel of root.childRelationships) {
    relationships.set(rel.relationshipName.toLowerCase(), rel.relationshipName);
    if (rel.relationshipName === "Approvals") {
      const scope = scopeFromFields(rel.label, APPROVALS_LOOP_FIELDS);
      scope.keys.add("today");
      scope.keys.add("now");
      childScopes.set("approvals", scope);
      continue;
    }
    const child = childDiscovers.get(rel.childObjectApiName);
    if (!child) continue; // not fetched yet — loop lints leniently
    const inLoop = child.rootScalarMergeFields.map((f) => ({
      ...f,
      key: inLoopFieldKey(f.key),
    }));
    const scope = scopeFromFields(rel.label, inLoop);
    scope.keys.add("today");
    scope.keys.add("now");
    childScopes.set(rel.relationshipName.toLowerCase(), scope);
  }
  return { root: rootScope, childScopes, relationships };
}

export interface LintEntry {
  tag: string;
  occurrence: number;
  status: "Resolved" | "Unresolved" | "Structural";
  suggestion?: string;
  /** Human-readable problem for this tag, when there is one. */
  note?: string;
}

export interface LintOutcome {
  entries: LintEntry[];
  /** Document-level structural problems (unbalanced blocks etc.). */
  errors: string[];
  hasProblems: boolean;
}

/** Scope-aware lint of the whole document text. */
export function lintDocumentText(text: string, ctx: LintContext): LintOutcome {
  const entries: LintEntry[] = [];
  const errors: string[] = [];
  const blockStack: string[] = [];
  const scopeStack: (LintScope | null)[] = []; // null = unknown child (lenient)

  const activeScope = (): LintScope | null =>
    scopeStack.length > 0 ? scopeStack[scopeStack.length - 1]! : ctx.root;

  const resolveKey = (tagText: string, occurrence: number, key: string): LintEntry => {
    const scope = activeScope();
    if (scope === null || scope.keys.has(key.toLowerCase())) {
      return { tag: tagText, occurrence, status: "Resolved" };
    }
    const suggestion = closestKey(key, scope.keys, scope.proper);
    return {
      tag: tagText,
      occurrence,
      status: "Unresolved",
      suggestion,
      note: `Not a field of ${scope.label}.`,
    };
  };

  for (const found of findTags(text)) {
    const p = found.parsed;
    let entry: LintEntry = { tag: found.tag, occurrence: found.occurrence, status: "Structural" };

    if (p.kind === "malformed") {
      entry = {
        tag: found.tag,
        occurrence: found.occurrence,
        status: "Unresolved",
        note: "Unrecognized merge tag.",
      };
    } else if (p.kind === "ifOpen") {
      blockStack.push("if");
    } else if (p.kind === "elseMarker") {
      if (blockStack[blockStack.length - 1] !== "if") {
        entry.note = "{{:else}} is not inside an {{#if}} block.";
        errors.push(entry.note);
      }
    } else if (p.kind === "ifClose") {
      if (blockStack[blockStack.length - 1] !== "if") {
        entry.note = "{{/if}} has no matching {{#if}}.";
        errors.push(entry.note);
      } else {
        blockStack.pop();
      }
    } else if (p.kind === "inverseOpen") {
      blockStack.push(`inv:${p.key}`);
    } else if (p.kind === "loopOpen") {
      const relKey = p.relationship.toLowerCase();
      if (scopeStack.length >= 2) {
        entry.status = "Unresolved";
        entry.note = "Only one level of nested lists is supported.";
        errors.push(`${found.tag}: ${entry.note}`);
        scopeStack.push(null);
      } else if (scopeStack.length === 0 && !ctx.relationships.has(relKey)) {
        entry.status = "Unresolved";
        entry.suggestion = closestKey(
          p.relationship,
          new Set(ctx.relationships.keys()),
          ctx.relationships,
        );
        entry.note = `"${p.relationship}" is not a related list of ${ctx.root.label}.`;
        scopeStack.push(null);
      } else {
        scopeStack.push(ctx.childScopes.get(relKey) ?? null);
      }
      if (p.where && !p.where.includes(`${p.relationship}.`)) {
        entry.note = `Filter fields need the "${p.relationship}." prefix (e.g. ${p.relationship}.Amount > 100).`;
        errors.push(`${found.tag}: ${entry.note}`);
      }
      if (p.orderBy) {
        const scope = scopeStack[scopeStack.length - 1];
        const bare = p.orderBy.startsWith(`${p.relationship}.`)
          ? p.orderBy.slice(p.relationship.length + 1)
          : p.orderBy;
        if (scope && !scope.keys.has(bare.toLowerCase())) {
          entry.note = `Sort field "${p.orderBy}" is not a field of ${scope.label}.`;
          errors.push(`${found.tag}: ${entry.note}`);
        }
      }
      blockStack.push(`loop:${p.relationship}`);
    } else if (p.kind === "blockClose") {
      const top = blockStack[blockStack.length - 1];
      if (!top) {
        entry.note = `${found.tag} has no matching open tag.`;
        errors.push(entry.note);
      } else {
        const expected = top.startsWith("loop:")
          ? top.slice(5)
          : top.startsWith("inv:")
            ? top.slice(4)
            : top;
        if (expected !== p.target) {
          entry.note = `${found.tag} closes "${p.target}" but the open block is "${expected}".`;
          errors.push(entry.note);
        }
        blockStack.pop();
        if (top.startsWith("loop:") && scopeStack.length > 0) {
          scopeStack.pop();
        }
      }
    } else if (p.kind === "aggregate") {
      if (scopeStack.length > 0) {
        entry.status = "Unresolved";
        entry.note = "Totals (SUM/COUNT/…) go outside the repeating list.";
        errors.push(`${found.tag}: ${entry.note}`);
      } else if (!ctx.relationships.has(p.relationship.toLowerCase())) {
        entry.status = "Unresolved";
        entry.note = `"${p.relationship}" is not a related list of ${ctx.root.label}.`;
      } else {
        const scope = ctx.childScopes.get(p.relationship.toLowerCase());
        if (p.field && scope && !scope.keys.has(p.field.toLowerCase())) {
          entry.status = "Unresolved";
          entry.suggestion = closestKey(p.field, scope.keys, scope.proper);
          entry.note = `"${p.field}" is not a field of ${scope.label}.`;
        } else {
          entry.status = "Resolved";
        }
      }
    } else if (p.kind === "scalar" || p.kind === "image" || p.kind === "barcode") {
      entry = resolveKey(found.tag, found.occurrence, p.key);
    }
    // signature stays Structural.
    entries.push(entry);
  }

  if (blockStack.length > 0) {
    const top = blockStack[blockStack.length - 1]!;
    const name = top === "if" ? "{{#if}}" : `{{#${top.slice(top.indexOf(":") + 1)}}}`;
    errors.push(`${name} is never closed.`);
  }

  const hasProblems = errors.length > 0 || entries.some((e) => e.status === "Unresolved");
  return { entries, errors, hasProblems };
}

/** Cheap did-you-mean: case-insensitive exact hit or ≥4-char shared prefix. */
export function closestKey(
  input: string,
  lowerKeys: Set<string>,
  proper: Map<string, string>,
): string | undefined {
  const lower = input.toLowerCase();
  if (lowerKeys.has(lower)) return proper.get(lower);
  let best: string | undefined;
  let bestLen = 3;
  for (const key of lowerKeys) {
    const shared = sharedPrefixLength(key, lower);
    if (shared > bestLen) {
      bestLen = shared;
      best = key;
    }
  }
  return best === undefined ? undefined : proper.get(best);
}

function sharedPrefixLength(a: string, b: string): number {
  const max = Math.min(a.length, b.length);
  let i = 0;
  while (i < max && a[i] === b[i]) i++;
  return i;
}
