import { describe, expect, it } from "vitest";
import {
  conditionalTags,
  defaultFormatForType,
  inLoopFieldKey,
  inverseTags,
  loopRowCellTexts,
  loopTags,
  scalarTag,
  truthyTags,
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
});

describe("truthy / inverse / loop tags", () => {
  it("builds the three block forms", () => {
    expect(truthyTags("Account.IsActive__c")).toEqual({
      open: "{{#Account.IsActive__c}}",
      close: "{{/Account.IsActive__c}}",
    });
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
    expect(defaultFormatForType("picklist")).toBeUndefined();
  });
});
