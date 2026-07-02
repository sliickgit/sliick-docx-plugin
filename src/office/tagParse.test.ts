import { describe, expect, it } from "vitest";
import {
  barcodeTag,
  compoundConditionTags,
  imageTag,
  loopTags,
  scalarTag,
  scalarTagWithOptions,
} from "./tags";
import { enclosingLoopAt, findTags, parseTag } from "./tagParse";

const inner = (tag: string): string => tag.slice(2, -2);

describe("parseTag round-trips builder output", () => {
  it("scalars with format, locale, and fallback", () => {
    expect(parseTag(inner(scalarTag("Account.Name")))).toEqual({
      kind: "scalar",
      key: "Account.Name",
      format: undefined,
      locale: undefined,
      fallback: undefined,
    });
    expect(
      parseTag(
        inner(
          scalarTagWithOptions("Account.AnnualRevenue", {
            format: "currency",
            locale: "de_DE",
            fallback: "N/A",
          }),
        ),
      ),
    ).toEqual({
      kind: "scalar",
      key: "Account.AnnualRevenue",
      format: "currency",
      locale: "de_DE",
      fallback: "N/A",
    });
    // Date PATTERN with colons must not be misread as locale.
    const patterned = parseTag("Opportunity.CloseDate:MM/dd/yyyy h:mm a");
    expect(patterned).toMatchObject({
      kind: "scalar",
      key: "Opportunity.CloseDate",
      format: "MM/dd/yyyy h:mm a",
    });
  });

  it("loop opens with WHERE / ORDER BY", () => {
    expect(parseTag(inner(loopTags("Contacts").open))).toEqual({
      kind: "loopOpen",
      relationship: "Contacts",
      where: undefined,
      orderBy: undefined,
      descending: undefined,
    });
    const full = loopTags("Opportunities", {
      where: "Opportunities.StageName = 'Closed Won'",
      orderBy: "CloseDate",
      descending: true,
    });
    expect(parseTag(inner(full.open))).toEqual({
      kind: "loopOpen",
      relationship: "Opportunities",
      where: "Opportunities.StageName = 'Closed Won'",
      orderBy: "CloseDate",
      descending: true,
    });
    expect(parseTag(inner(full.close))).toEqual({ kind: "blockClose", target: "Opportunities" });
    // Quoted keywords stay data.
    expect(
      parseTag("#Cases WHERE Cases.Subject contains 'ORDER BY mail'"),
    ).toMatchObject({
      kind: "loopOpen",
      relationship: "Cases",
      where: "Cases.Subject contains 'ORDER BY mail'",
    });
  });

  it("barcodes, images, aggregates, structural forms", () => {
    expect(parseTag(inner(barcodeTag("Account.Website", "qr", "150")))).toEqual({
      kind: "barcode",
      key: "Account.Website",
      barcodeType: "qr",
      size: "150",
    });
    expect(parseTag(inner(imageTag("Account.Logo__c", 200, 60)))).toEqual({
      kind: "image",
      key: "Account.Logo__c",
      size: "200x60",
    });
    expect(parseTag("SUM:OpportunityLineItems.TotalPrice:currency")).toEqual({
      kind: "aggregate",
      fn: "SUM",
      relationship: "OpportunityLineItems",
      field: "TotalPrice",
      format: "currency",
    });
    expect(parseTag("COUNT:Contacts")).toMatchObject({ kind: "aggregate", fn: "COUNT" });
    const cond = compoundConditionTags(
      [{ fieldKey: "Account.Industry", operator: "contains", value: "Tech", quoteValue: true }],
      "AND",
      true,
    );
    expect(parseTag(inner(cond.open))).toEqual({
      kind: "ifOpen",
      expression: "Account.Industry contains 'Tech'",
    });
    expect(parseTag(":else")).toEqual({ kind: "elseMarker" });
    expect(parseTag("/if")).toEqual({ kind: "ifClose" });
    expect(parseTag("^Account.Description")).toEqual({
      kind: "inverseOpen",
      key: "Account.Description",
    });
    expect(parseTag("@Signature:Buyer:1:Full")).toMatchObject({ kind: "signature" });
  });

  it("flags malformed shapes", () => {
    expect(parseTag("").kind).toBe("malformed");
    expect(parseTag("#Contacts.Email").kind).toBe("malformed");
    expect(parseTag("#Contacts WHERE ").kind).toBe("malformed");
    expect(parseTag("Hello world").kind).toBe("malformed");
    expect(parseTag("*Account.Website:hologram").kind).toBe("malformed");
    expect(parseTag("COUNT:Contacts.Name").kind).toBe("malformed");
  });
});

describe("findTags", () => {
  it("returns document-order tags with occurrence indices", () => {
    const text = "A {{Account.Name}} B {{Today}} C {{Account.Name}}";
    const tags = findTags(text);
    expect(tags.map((t) => t.tag)).toEqual([
      "{{Account.Name}}",
      "{{Today}}",
      "{{Account.Name}}",
    ]);
    expect(tags[0]!.occurrence).toBe(0);
    expect(tags[2]!.occurrence).toBe(1);
    expect(tags[1]!.index).toBe(text.indexOf("{{Today}}"));
  });
});

describe("enclosingLoopAt", () => {
  const text =
    "intro {{#Contacts}} inside {{#Cases}} deep {{/Cases}} after {{/Contacts}} outro";

  it("detects the innermost unclosed loop at a position", () => {
    expect(enclosingLoopAt(text, text.indexOf("intro"))).toBeNull();
    expect(enclosingLoopAt(text, text.indexOf("inside"))).toBe("Contacts");
    expect(enclosingLoopAt(text, text.indexOf("deep"))).toBe("Cases");
    expect(enclosingLoopAt(text, text.indexOf("after"))).toBe("Contacts");
    expect(enclosingLoopAt(text, text.indexOf("outro"))).toBeNull();
  });

  it("ignores non-loop blocks and inverse closes", () => {
    const mixed = "{{#if Account.X = 1}} {{^Account.Y}} here {{/Account.Y}} {{/if}}";
    expect(enclosingLoopAt(mixed, mixed.indexOf("here"))).toBeNull();
  });
});
