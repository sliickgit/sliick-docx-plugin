/**
 * Tag parser — the inverse of the builders in tags.ts. One grammar-aware
 * parser (mirroring the engine's DocxTagEvaluator.classify) feeding local
 * lint, the tag navigator, edit-in-wizard prefill, and cursor-scope
 * detection. Pure TS, no Office dependencies.
 */

import { AggregateFn, LOCALE_KEYWORD_FORMATS } from "./tags";

export type ParsedTag =
  | { kind: "scalar"; key: string; format?: string; locale?: string; fallback?: string }
  | {
      kind: "loopOpen";
      relationship: string;
      where?: string;
      orderBy?: string;
      descending?: boolean;
    }
  | { kind: "blockClose"; target: string }
  | { kind: "ifOpen"; expression: string }
  | { kind: "elseMarker" }
  | { kind: "ifClose" }
  | { kind: "inverseOpen"; key: string }
  | { kind: "aggregate"; fn: AggregateFn; relationship: string; field?: string; format?: string }
  | { kind: "image"; key: string; size?: string }
  | { kind: "barcode"; key: string; barcodeType: "code128" | "qr"; size?: string }
  | { kind: "signature"; raw: string }
  | { kind: "malformed"; raw: string };

const PATHISH = /^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)*$/;
const LOCALE_SHAPE = /^[A-Za-z]{2}[-_][A-Za-z]{2}$/;
const AGGREGATE_FUNCTIONS = new Set(["SUM", "COUNT", "AVG", "MIN", "MAX"]);

function isPathish(s: string): boolean {
  return PATHISH.test(s);
}

/** Parses the inner text of one {{...}} tag (braces already stripped). */
export function parseTag(innerText: string): ParsedTag {
  const trimmed = (innerText ?? "").trim();
  if (!trimmed) return { kind: "malformed", raw: trimmed };

  if (trimmed === "#if" || trimmed.startsWith("#if ") || trimmed.startsWith("#if\t")) {
    const expression = trimmed.slice(3).trim();
    return expression ? { kind: "ifOpen", expression } : { kind: "malformed", raw: trimmed };
  }
  if (trimmed === ":else") return { kind: "elseMarker" };
  if (trimmed === "/if") return { kind: "ifClose" };
  if (trimmed.startsWith("^")) {
    const key = trimmed.slice(1).trim();
    return isPathish(key) ? { kind: "inverseOpen", key } : { kind: "malformed", raw: trimmed };
  }
  if (trimmed.startsWith("/")) {
    const target = trimmed.slice(1).trim();
    return target ? { kind: "blockClose", target } : { kind: "malformed", raw: trimmed };
  }
  if (trimmed.startsWith("#")) {
    return parseLoopOpen(trimmed.slice(1).trim());
  }
  if (trimmed.startsWith("%")) {
    const rest = trimmed.slice(1).trim();
    const colonAt = rest.indexOf(":");
    const key = colonAt === -1 ? rest : rest.slice(0, colonAt).trim();
    const size = colonAt === -1 ? undefined : rest.slice(colonAt + 1).trim();
    return isPathish(key) ? { kind: "image", key, size } : { kind: "malformed", raw: trimmed };
  }
  if (trimmed.startsWith("@")) {
    return { kind: "signature", raw: trimmed };
  }
  if (trimmed.startsWith("*")) {
    return parseBarcode(trimmed.slice(1).trim(), trimmed);
  }

  const funcColon = trimmed.indexOf(":");
  if (funcColon > 0 && AGGREGATE_FUNCTIONS.has(trimmed.slice(0, funcColon).trim().toUpperCase())) {
    return parseAggregate(trimmed, funcColon);
  }
  return parseScalar(trimmed);
}

function parseLoopOpen(body: string): ParsedTag {
  let remainder = body;
  let orderBy: string | undefined;
  let descending: boolean | undefined;
  let where: string | undefined;

  const orderAt = indexOfKeywordOutsideQuotes(remainder, " ORDER BY ");
  if (orderAt !== -1) {
    let orderPart = remainder.slice(orderAt + 10).trim();
    remainder = remainder.slice(0, orderAt).trim();
    const upper = orderPart.toUpperCase();
    if (upper.endsWith(" DESC")) {
      descending = true;
      orderPart = orderPart.slice(0, -5).trim();
    } else if (upper.endsWith(" ASC")) {
      orderPart = orderPart.slice(0, -4).trim();
    }
    orderBy = orderPart;
    if (!isPathish(orderPart)) return { kind: "malformed", raw: `#${body}` };
  }
  const whereAt = indexOfKeywordOutsideQuotes(remainder, " WHERE ");
  if (whereAt !== -1) {
    where = remainder.slice(whereAt + 7).trim();
    remainder = remainder.slice(0, whereAt).trim();
    if (!where) return { kind: "malformed", raw: `#${body}` };
  }
  if (!remainder || remainder.includes(".") || !isPathish(remainder)) {
    return { kind: "malformed", raw: `#${body}` };
  }
  return { kind: "loopOpen", relationship: remainder, where, orderBy, descending };
}

function parseBarcode(body: string, raw: string): ParsedTag {
  const parts = body.split(":");
  const key = (parts[0] ?? "").trim();
  let barcodeType: "code128" | "qr" = "code128";
  let size: string | undefined;
  for (const rawPart of parts.slice(1)) {
    const part = rawPart.trim().toLowerCase();
    if (part === "qr" || part === "code128") barcodeType = part;
    else if (/^\d+(x\d+)?$/.test(part)) size = part;
    else return { kind: "malformed", raw };
  }
  return isPathish(key) ? { kind: "barcode", key, barcodeType, size } : { kind: "malformed", raw };
}

function parseAggregate(trimmed: string, funcColon: number): ParsedTag {
  const fn = trimmed.slice(0, funcColon).trim().toUpperCase() as AggregateFn;
  const rest = trimmed.slice(funcColon + 1).trim();
  const fmtColon = rest.indexOf(":");
  const relField = fmtColon === -1 ? rest : rest.slice(0, fmtColon).trim();
  const format = fmtColon === -1 ? undefined : rest.slice(fmtColon + 1).trim();
  const dot = relField.indexOf(".");
  const relationship = dot === -1 ? relField : relField.slice(0, dot).trim();
  const field = dot === -1 ? undefined : relField.slice(dot + 1).trim();
  const relOk = isPathish(relationship);
  const fieldOk = fn === "COUNT" ? !field : !!field && isPathish(field);
  return relOk && fieldOk
    ? { kind: "aggregate", fn, relationship, field, format }
    : { kind: "malformed", raw: trimmed };
}

function parseScalar(trimmed: string): ParsedTag {
  let scalarPart = trimmed;
  let fallback: string | undefined;
  const pipeAt = trimmed.indexOf("|");
  if (pipeAt !== -1) {
    fallback = trimmed.slice(pipeAt + 1);
    scalarPart = trimmed.slice(0, pipeAt).trim();
  }
  const colonAt = scalarPart.indexOf(":");
  const key = colonAt === -1 ? scalarPart : scalarPart.slice(0, colonAt).trim();
  let format = colonAt === -1 ? undefined : scalarPart.slice(colonAt + 1).trim();
  let locale: string | undefined;
  if (format) {
    // 'currency:de_DE' → keyword format + locale override.
    const localeColon = format.indexOf(":");
    if (localeColon > 0) {
      const keywordPart = format.slice(0, localeColon).trim().toLowerCase();
      const localePart = format.slice(localeColon + 1).trim();
      if (LOCALE_KEYWORD_FORMATS.has(keywordPart) && LOCALE_SHAPE.test(localePart)) {
        format = keywordPart;
        locale = localePart;
      }
    }
  }
  return isPathish(key)
    ? { kind: "scalar", key, format, locale, fallback }
    : { kind: "malformed", raw: trimmed };
}

/**
 * Case-insensitive keyword scan skipping single-quoted literals (with \' and
 * \\ escapes) — mirrors the engine's parser so quoted "ORDER BY" is data.
 */
export function indexOfKeywordOutsideQuotes(hay: string, keyword: string): number {
  const upperHay = hay.toUpperCase();
  const upperKeyword = keyword.toUpperCase();
  let inQuote = false;
  let i = 0;
  while (i < hay.length) {
    const c = hay[i];
    if (inQuote) {
      if (c === "\\") {
        i += 2;
        continue;
      }
      if (c === "'") inQuote = false;
      i += 1;
      continue;
    }
    if (c === "'") {
      inQuote = true;
      i += 1;
      continue;
    }
    if (upperHay.startsWith(upperKeyword, i)) return i;
    i += 1;
  }
  return -1;
}

export interface FoundTag {
  /** Full tag text including braces. */
  tag: string;
  /** Inner text (braces stripped, not trimmed). */
  inner: string;
  /** Character index of '{{' in the scanned text. */
  index: number;
  /** 0-based occurrence index among identical tag texts (for navigation). */
  occurrence: number;
  parsed: ParsedTag;
}

/** Every {{...}} in a text, in document order, with duplicate-occurrence indices. */
export function findTags(text: string): FoundTag[] {
  const out: FoundTag[] = [];
  const seen = new Map<string, number>();
  const re = /\{\{([^{}]+)\}\}/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const tag = m[0]!;
    const occurrence = seen.get(tag) ?? 0;
    seen.set(tag, occurrence + 1);
    out.push({ tag, inner: m[1]!, index: m.index, occurrence, parsed: parseTag(m[1]!) });
  }
  return out;
}

/**
 * The unclosed loop enclosing a character position, or null at root scope —
 * powers the pane's scope-aware field list. Walks tags before `position`,
 * tracking the loop stack the same way the engine's validator does.
 */
export function enclosingLoopAt(text: string, position: number): string | null {
  const stack: string[] = [];
  for (const found of findTags(text)) {
    if (found.index >= position) break;
    if (found.parsed.kind === "loopOpen") {
      stack.push(found.parsed.relationship);
    } else if (found.parsed.kind === "blockClose") {
      if (stack.length > 0 && stack[stack.length - 1] === found.parsed.target) {
        stack.pop();
      }
    }
  }
  return stack.length > 0 ? stack[stack.length - 1]! : null;
}
