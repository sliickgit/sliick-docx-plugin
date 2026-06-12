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

/**
 * Default format suffix for a Salesforce field type, or undefined when raw
 * output is right. Keys are lowercase Salesforce display types as delivered
 * by the discover endpoint.
 */
export function defaultFormatForType(sfType: string): string | undefined {
  switch (sfType.toLowerCase()) {
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
