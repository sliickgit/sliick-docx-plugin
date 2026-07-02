/**
 * Pure tag-string builders for the sliick {{...}} grammar (grammarVersion 1).
 * Mirrors sliick-docs Phase H §3.7. No Office dependencies — unit-testable.
 */

export function scalarTag(key: string, format?: string): string {
  return format ? `{{${key}:${format}}}` : `{{${key}}}`;
}

export type ConditionOperator = "=" | "!=" | ">" | "<" | ">=" | "<=" | "contains";

export interface ConditionalSpec {
  fieldKey: string;
  operator: ConditionOperator;
  value: string;
  /** Quote string values; numbers/booleans pass through bare. */
  quoteValue: boolean;
  /** Invert the comparison — emitted as `NOT (clause)`. */
  negate?: boolean;
  withElse: boolean;
}

/**
 * One clause as expression text. String literals are SINGLE-quoted — the
 * sliick-docs expression lexer (TemplateExpressionService) rejects double
 * quotes. Embedded single quotes are stripped; the engine has no escape
 * syntax for them. `contains` compares stringified values, so its literal is
 * always quoted. Negation is parenthesized so parsing is unambiguous.
 */
function clauseText(c: ConditionClause): string {
  const quote = c.quoteValue || c.operator === "contains";
  const v = quote ? `'${c.value.replace(/'/g, "")}'` : c.value;
  const body = `${c.fieldKey} ${c.operator} ${v}`;
  return c.negate ? `NOT (${body})` : body;
}

export function conditionalTags(spec: ConditionalSpec): {
  open: string;
  elseTag?: string;
  close: string;
} {
  return {
    open: `{{#if ${clauseText(spec)}}}`,
    elseTag: spec.withElse ? "{{:else}}" : undefined,
    close: "{{/if}}",
  };
}

// NOTE: there is deliberately no truthy-block builder — the engine's
// DocxTagEvaluator treats `{{#X}}` strictly as a loop-open over a child
// relationship, so `{{#Field}}` would fail validation. Use
// `{{#if Field != null}}` (conditionalTags) for "show when present".

export function inverseTags(fieldKey: string): { open: string; close: string } {
  return { open: `{{^${fieldKey}}}`, close: `{{/${fieldKey}}}` };
}

/** grammar-v2 loop modifiers: {{#Rel WHERE <expr> ORDER BY <Field> [DESC]}}. */
export interface LoopModifiers {
  /** Filter expression text — clauses must carry the `Rel.` prefix. */
  where?: string;
  /** Sort field (child-relative in-loop key). */
  orderBy?: string;
  descending?: boolean;
}

export function loopTags(
  relationshipName: string,
  modifiers?: LoopModifiers,
): { open: string; close: string } {
  let open = `{{#${relationshipName}`;
  if (modifiers?.where) open += ` WHERE ${modifiers.where}`;
  if (modifiers?.orderBy) {
    open += ` ORDER BY ${modifiers.orderBy}${modifiers.descending ? " DESC" : ""}`;
  }
  open += "}}";
  return { open, close: `{{/${relationshipName}}}` };
}

/**
 * Expression text for one-or-two clauses joined by AND/OR — shared by the
 * condition wizard ({{#if}}) and the loop-filter builder (WHERE).
 */
export function conditionExpressionText(
  clauses: ConditionClause[],
  connector: "AND" | "OR",
): string {
  if (clauses.length === 0) throw new Error("A condition needs at least one clause.");
  return clauses.map(clauseText).join(` ${connector} `);
}

/** Barcode tag: {{*Field}}, {{*Field:qr}}, {{*Field:qr:150}}, {{*Field:code128:250x80}}. */
export function barcodeTag(
  fieldKey: string,
  type?: "code128" | "qr",
  size?: string,
): string {
  let body = `*${fieldKey}`;
  if (type) body += `:${type}`;
  if (size) body += `:${size}`;
  return `{{${body}}}`;
}

/** Keyword formats that accept a locale override suffix ({{X:date:de_DE}}). */
export const LOCALE_KEYWORD_FORMATS = new Set([
  "currency",
  "number",
  "percent",
  "date",
  "datetime",
]);

/**
 * Scalar tag with grammar-v2 options: {{Key:format:locale|fallback}}. A
 * locale is only emitted alongside a keyword format (the engine reads any
 * other suffix as a date pattern).
 */
export function scalarTagWithOptions(
  key: string,
  opts: { format?: string; locale?: string; fallback?: string },
): string {
  let body = key;
  if (opts.format) {
    body += `:${opts.format}`;
    if (opts.locale && LOCALE_KEYWORD_FORMATS.has(opts.format)) {
      body += `:${opts.locale}`;
    }
  }
  if (opts.fallback !== undefined && opts.fallback !== "") {
    body += `|${opts.fallback}`;
  }
  return `{{${body}}}`;
}

export type AggregateFn = "SUM" | "COUNT" | "AVG" | "MIN" | "MAX";

/**
 * Aggregate tag: {{COUNT:Rel}} or {{SUM:Rel.Field}} with optional format.
 * Field is omitted for COUNT (and ignored if passed).
 */
export function aggregateTag(
  fn: AggregateFn,
  relationshipName: string,
  fieldApiName?: string,
  format?: string,
): string {
  const body =
    fn === "COUNT" ? `COUNT:${relationshipName}` : `${fn}:${relationshipName}.${fieldApiName}`;
  return format ? `{{${body}:${format}}}` : `{{${body}}}`;
}

/** Image merge field: {{%Field}} or {{%Field:WxH}} (pixels). */
export function imageTag(fieldKey: string, widthPx?: number, heightPx?: number): string {
  if (widthPx && heightPx) {
    return `{{%${fieldKey}:${widthPx}x${heightPx}}}`;
  }
  return `{{%${fieldKey}}}`;
}

/** One comparison clause of a compound condition. */
export interface ConditionClause {
  fieldKey: string;
  operator: ConditionOperator;
  value: string;
  quoteValue: boolean;
  /** Invert this clause — emitted as `NOT (clause)`. */
  negate?: boolean;
}

/**
 * Compound conditional: clauses joined by a single AND/OR connector.
 * Engine supports nested AND/OR/NOT; the wizard exposes one connector for
 * simplicity. Returns the same open/elseTag/close shape as conditionalTags.
 */
export function compoundConditionTags(
  clauses: ConditionClause[],
  connector: "AND" | "OR",
  withElse: boolean,
): { open: string; elseTag?: string; close: string } {
  return {
    open: `{{#if ${conditionExpressionText(clauses, connector)}}}`,
    elseTag: withElse ? "{{:else}}" : undefined,
    close: "{{/if}}",
  };
}

/**
 * Default format suffix for a Salesforce field type, or undefined when raw
 * output is right. Keys are lowercase Salesforce display types as delivered
 * by the discover endpoint.
 */
export function defaultFormatForType(sfType?: string): string | undefined {
  switch ((sfType ?? "").toLowerCase()) {
    case "currency":
      return "currency";
    case "percent":
      return "percent";
    case "date":
      return "MM/dd/yyyy";
    case "datetime":
      return "MM/dd/yyyy h:mm a";
    case "boolean":
      return "checkbox";
    case "picklist":
    case "multipicklist":
      // Render the user-facing label, not the stored API value.
      return "label";
    default:
      return undefined;
  }
}

/**
 * Loop-table cell texts for a row-scope child loop per Phase H §3.7.1:
 * `{{#Rel}}` opens in the FIRST cell, `{{/Rel}}` closes in the LAST cell of
 * the same row. Returns one cell text per column.
 */
export function loopRowCellTexts(
  relationshipName: string,
  fieldKeys: string[],
  modifiers?: LoopModifiers,
): string[] {
  if (fieldKeys.length === 0) throw new Error("Loop table needs at least one column.");
  const { open, close } = loopTags(relationshipName, modifiers);
  return fieldKeys.map((key, i) => {
    let text = `{{${key}}}`;
    if (i === 0) text = `${open}${text}`;
    if (i === fieldKeys.length - 1) text = `${text}${close}`;
    return text;
  });
}

/**
 * Inside a row-scope loop the engine resolves child fields WITHOUT the base
 * object prefix (`{{FirstName}}` inside `{{#Contacts}}`). Converts a child
 * field key like "Contact.FirstName" to its in-loop form.
 */
export function inLoopFieldKey(childFieldKey: string): string {
  const dot = childFieldKey.indexOf(".");
  return dot === -1 ? childFieldKey : childFieldKey.slice(dot + 1);
}

/**
 * Paragraph-scope nested loop (engine max: one nesting level). Open/close
 * tags sit on their own paragraphs so OfficeLoopScope detects paragraph
 * scope for both levels; `childLine`/`nestedLine` are pre-built tag texts
 * (in-loop keys, formats already applied).
 */
export function nestedLoopBlockLines(
  relationshipName: string,
  childLine: string,
  nestedRelationshipName: string,
  nestedLine: string,
): string[] {
  if (!childLine || !nestedLine) {
    throw new Error("Nested loop needs at least one field at each level.");
  }
  const outer = loopTags(relationshipName);
  const inner = loopTags(nestedRelationshipName);
  return [outer.open, childLine, inner.open, nestedLine, inner.close, outer.close];
}
