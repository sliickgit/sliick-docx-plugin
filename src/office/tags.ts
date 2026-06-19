/**
 * Pure tag-string builders for the sliick {{...}} grammar (grammarVersion 1).
 * Mirrors sliick-docs Phase H §3.7. No Office dependencies — unit-testable.
 */

export function scalarTag(key: string, format?: string): string {
  return format ? `{{${key}:${format}}}` : `{{${key}}}`;
}

export interface ConditionalSpec {
  fieldKey: string;
  operator: "=" | "!=" | ">" | "<" | ">=" | "<=";
  value: string;
  /** Quote string values; numbers/booleans pass through bare. */
  quoteValue: boolean;
  withElse: boolean;
}

export function conditionalTags(spec: ConditionalSpec): {
  open: string;
  elseTag?: string;
  close: string;
} {
  // String literals are SINGLE-quoted — the sliick-docs expression lexer
  // (TemplateExpressionService) rejects double quotes. Embedded single
  // quotes are stripped; the engine has no escape syntax for them.
  const value = spec.quoteValue ? `'${spec.value.replace(/'/g, "")}'` : spec.value;
  return {
    open: `{{#if ${spec.fieldKey} ${spec.operator} ${value}}}`,
    elseTag: spec.withElse ? "{{:else}}" : undefined,
    close: "{{/if}}",
  };
}

export function truthyTags(fieldKey: string): { open: string; close: string } {
  return { open: `{{#${fieldKey}}}`, close: `{{/${fieldKey}}}` };
}

export function inverseTags(fieldKey: string): { open: string; close: string } {
  return { open: `{{^${fieldKey}}}`, close: `{{/${fieldKey}}}` };
}

export function loopTags(relationshipName: string): { open: string; close: string } {
  return { open: `{{#${relationshipName}}}`, close: `{{/${relationshipName}}}` };
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
  operator: "=" | "!=" | ">" | "<" | ">=" | "<=";
  value: string;
  quoteValue: boolean;
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
  if (clauses.length === 0) throw new Error("A condition needs at least one clause.");
  const expr = clauses
    .map((c) => {
      const v = c.quoteValue ? `'${c.value.replace(/'/g, "")}'` : c.value;
      return `${c.fieldKey} ${c.operator} ${v}`;
    })
    .join(` ${connector} `);
  return {
    open: `{{#if ${expr}}}`,
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
): string[] {
  if (fieldKeys.length === 0) throw new Error("Loop table needs at least one column.");
  const { open, close } = loopTags(relationshipName);
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
