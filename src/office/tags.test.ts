import { describe, expect, it } from "vitest";
import {
  aggregateTag,
  barcodeTag,
  compoundConditionTags,
  conditionalTags,
  conditionExpressionText,
  defaultFormatForType,
  imageTag,
  inLoopFieldKey,
  inverseTags,
  loopRowCellTexts,
  loopTags,
  nestedLoopBlockLines,
  scalarTag,
  scalarTagWithOptions,
} from "./tags";

describe("scalarTag", () => {
  it("renders bare and formatted tags", () => {
    expect(scalarTag("Account.Name")).toBe("{{Account.Name}}");
    expect(scalarTag("Opportunity.Amount", "currency")).toBe(
      "{{Opportunity.Amount:currency}}",
    );
  });
});

describe("conditionalTags", () => {
  it("single-quotes string values (engine lexer) and uses the {{:else}} token", () => {
    const t = conditionalTags({
      fieldKey: "Opportunity.StageName",
      operator: "=",
      value: "Closed Won",
      quoteValue: true,
      withElse: true,
    });
    expect(t.open).toBe("{{#if Opportunity.StageName = 'Closed Won'}}");
    expect(t.elseTag).toBe("{{:else}}");
    expect(t.close).toBe("{{/if}}");
  });

  it("strips embedded single quotes from string values", () => {
    const t = conditionalTags({
      fieldKey: "Account.Name",
      operator: "=",
      value: "O'Brien Co",
      quoteValue: true,
      withElse: false,
    });
    expect(t.open).toBe("{{#if Account.Name = 'OBrien Co'}}");
  });

  it("leaves numeric values unquoted and omits else when not requested", () => {
    const t = conditionalTags({
      fieldKey: "Opportunity.Amount",
      operator: ">",
      value: "50000",
      quoteValue: false,
      withElse: false,
    });
    expect(t.open).toBe("{{#if Opportunity.Amount > 50000}}");
    expect(t.elseTag).toBeUndefined();
  });

  it("always quotes contains values, even for numeric fields", () => {
    const t = conditionalTags({
      fieldKey: "Account.Description",
      operator: "contains",
      value: "priority",
      quoteValue: false,
      withElse: false,
    });
    expect(t.open).toBe("{{#if Account.Description contains 'priority'}}");
  });

  it("parenthesizes negated clauses", () => {
    const t = conditionalTags({
      fieldKey: "Opportunity.StageName",
      operator: "=",
      value: "Closed Won",
      quoteValue: true,
      negate: true,
      withElse: false,
    });
    expect(t.open).toBe("{{#if NOT (Opportunity.StageName = 'Closed Won')}}");
  });
});

describe("inverse / loop tags", () => {
  it("builds both block forms", () => {
    expect(inverseTags("Account.HasDiscount__c").open).toBe(
      "{{^Account.HasDiscount__c}}",
    );
    expect(loopTags("Contacts")).toEqual({
      open: "{{#Contacts}}",
      close: "{{/Contacts}}",
    });
  });
});

describe("loopRowCellTexts", () => {
  it("opens in the first cell and closes in the last (Phase H row scope)", () => {
    const cells = loopRowCellTexts("Contacts", ["FirstName", "LastName", "Email"]);
    expect(cells).toEqual([
      "{{#Contacts}}{{FirstName}}",
      "{{LastName}}",
      "{{Email}}{{/Contacts}}",
    ]);
  });

  it("handles a single-column table (open and close in the same cell)", () => {
    expect(loopRowCellTexts("Cases", ["Subject"])).toEqual([
      "{{#Cases}}{{Subject}}{{/Cases}}",
    ]);
  });

  it("rejects empty column lists", () => {
    expect(() => loopRowCellTexts("Contacts", [])).toThrow();
  });
});

describe("grammar-v2 builders", () => {
  it("loopTags emits WHERE and ORDER BY modifiers", () => {
    expect(loopTags("Opportunities").open).toBe("{{#Opportunities}}");
    expect(
      loopTags("Opportunities", { where: "Opportunities.Amount > 10000" }).open,
    ).toBe("{{#Opportunities WHERE Opportunities.Amount > 10000}}");
    expect(
      loopTags("Opportunities", { orderBy: "CloseDate", descending: true }).open,
    ).toBe("{{#Opportunities ORDER BY CloseDate DESC}}");
    expect(
      loopTags("Opportunities", {
        where: "Opportunities.StageName = 'Closed Won'",
        orderBy: "Amount",
      }).open,
    ).toBe("{{#Opportunities WHERE Opportunities.StageName = 'Closed Won' ORDER BY Amount}}");
    expect(loopTags("Opportunities", { where: "x" }).close).toBe("{{/Opportunities}}");
  });

  it("conditionExpressionText builds prefixed filter clauses", () => {
    expect(
      conditionExpressionText(
        [
          { fieldKey: "Opportunities.Amount", operator: ">", value: "100", quoteValue: false },
          { fieldKey: "Opportunities.StageName", operator: "=", value: "Won", quoteValue: true, negate: true },
        ],
        "AND",
      ),
    ).toBe("Opportunities.Amount > 100 AND NOT (Opportunities.StageName = 'Won')");
  });

  it("barcodeTag builds every size/type form", () => {
    expect(barcodeTag("Account.AccountNumber")).toBe("{{*Account.AccountNumber}}");
    expect(barcodeTag("Account.Website", "qr")).toBe("{{*Account.Website:qr}}");
    expect(barcodeTag("Account.Website", "qr", "150")).toBe("{{*Account.Website:qr:150}}");
    expect(barcodeTag("Account.AccountNumber", "code128", "250x80")).toBe(
      "{{*Account.AccountNumber:code128:250x80}}",
    );
  });

  it("scalarTagWithOptions composes format, locale, and fallback", () => {
    expect(scalarTagWithOptions("Account.Name", {})).toBe("{{Account.Name}}");
    expect(scalarTagWithOptions("Account.AnnualRevenue", { format: "currency" })).toBe(
      "{{Account.AnnualRevenue:currency}}",
    );
    expect(
      scalarTagWithOptions("Account.AnnualRevenue", { format: "currency", locale: "de_DE" }),
    ).toBe("{{Account.AnnualRevenue:currency:de_DE}}");
    expect(
      scalarTagWithOptions("Account.Description", { fallback: "No notes on file" }),
    ).toBe("{{Account.Description|No notes on file}}");
    expect(
      scalarTagWithOptions("Account.AnnualRevenue", {
        format: "currency",
        locale: "de_DE",
        fallback: "N/A",
      }),
    ).toBe("{{Account.AnnualRevenue:currency:de_DE|N/A}}");
    // A locale never rides on a date PATTERN (engine would misread it).
    expect(
      scalarTagWithOptions("Account.CreatedDate", { format: "MM/dd/yyyy", locale: "de_DE" }),
    ).toBe("{{Account.CreatedDate:MM/dd/yyyy}}");
  });

  it("loopRowCellTexts carries modifiers into the opening cell", () => {
    const cells = loopRowCellTexts("Contacts", ["FirstName", "Email"], {
      where: "Contacts.Title contains 'VP'",
    });
    expect(cells[0]).toBe("{{#Contacts WHERE Contacts.Title contains 'VP'}}{{FirstName}}");
    expect(cells[1]).toBe("{{Email}}{{/Contacts}}");
  });
});

describe("nestedLoopBlockLines", () => {
  it("builds a paragraph-scope block with one nesting level", () => {
    const lines = nestedLoopBlockLines(
      "Opportunities",
      "{{Name}} — {{Amount:currency}}",
      "OpportunityLineItems",
      "{{Quantity}} — {{TotalPrice:currency}}",
    );
    expect(lines).toEqual([
      "{{#Opportunities}}",
      "{{Name}} — {{Amount:currency}}",
      "{{#OpportunityLineItems}}",
      "{{Quantity}} — {{TotalPrice:currency}}",
      "{{/OpportunityLineItems}}",
      "{{/Opportunities}}",
    ]);
  });

  it("rejects empty field lines at either level", () => {
    expect(() => nestedLoopBlockLines("A", "", "B", "{{X}}")).toThrow();
    expect(() => nestedLoopBlockLines("A", "{{X}}", "B", "")).toThrow();
  });
});

describe("inLoopFieldKey", () => {
  it("strips the object prefix for in-loop resolution", () => {
    expect(inLoopFieldKey("Contact.FirstName")).toBe("FirstName");
    expect(inLoopFieldKey("AlreadyBare")).toBe("AlreadyBare");
  });
});

describe("defaultFormatForType", () => {
  it("maps Salesforce types to grammar format suffixes", () => {
    expect(defaultFormatForType("currency")).toBe("currency");
    expect(defaultFormatForType("percent")).toBe("percent");
    expect(defaultFormatForType("boolean")).toBe("checkbox");
    expect(defaultFormatForType("date")).toBe("MM/dd/yyyy");
    expect(defaultFormatForType("string")).toBeUndefined();
    // Picklists auto-apply :label so the doc shows the label, not the API value.
    expect(defaultFormatForType("picklist")).toBe("label");
    expect(defaultFormatForType("multipicklist")).toBe("label");
  });
});

describe("aggregateTag", () => {
  it("builds COUNT without a field and SUM/AVG with a field + format", () => {
    expect(aggregateTag("COUNT", "Contacts")).toBe("{{COUNT:Contacts}}");
    expect(aggregateTag("SUM", "OpportunityLineItems", "TotalPrice", "currency")).toBe(
      "{{SUM:OpportunityLineItems.TotalPrice:currency}}",
    );
    expect(aggregateTag("AVG", "OpportunityLineItems", "UnitPrice")).toBe(
      "{{AVG:OpportunityLineItems.UnitPrice}}",
    );
  });
});

describe("imageTag", () => {
  it("builds bare and sized image tags", () => {
    expect(imageTag("Account.Logo__c")).toBe("{{%Account.Logo__c}}");
    expect(imageTag("Account.Logo__c", 200, 60)).toBe("{{%Account.Logo__c:200x60}}");
  });
});

describe("compoundConditionTags", () => {
  it("joins clauses with AND/OR and single-quotes string values", () => {
    const t = compoundConditionTags(
      [
        { fieldKey: "Account.AnnualRevenue", operator: ">", value: "50000", quoteValue: false },
        { fieldKey: "Account.Industry", operator: "=", value: "Technology", quoteValue: true },
      ],
      "AND",
      true,
    );
    expect(t.open).toBe(
      "{{#if Account.AnnualRevenue > 50000 AND Account.Industry = 'Technology'}}",
    );
    expect(t.elseTag).toBe("{{:else}}");
    expect(t.close).toBe("{{/if}}");
  });

  it("rejects an empty clause list", () => {
    expect(() => compoundConditionTags([], "AND", false)).toThrow();
  });

  it("negates individual clauses inside a compound expression", () => {
    const t = compoundConditionTags(
      [
        { fieldKey: "Account.Industry", operator: "contains", value: "Tech", quoteValue: true },
        { fieldKey: "Account.AnnualRevenue", operator: "<", value: "1000", quoteValue: false, negate: true },
      ],
      "OR",
      false,
    );
    expect(t.open).toBe(
      "{{#if Account.Industry contains 'Tech' OR NOT (Account.AnnualRevenue < 1000)}}",
    );
  });
});
