import { describe, expect, it } from "vitest";
import {
  classifyDocumentTags,
  classifyTag,
  extractTags,
  MockSliickClient,
} from "./mock";

const RESOLVABLE = new Set([
  "Account.Name",
  "Account.Industry",
  "Opportunity.Amount",
  "Today",
]);

describe("extractTags", () => {
  it("finds every {{...}} occurrence", () => {
    const text =
      "Dear {{Account.Name}}, on {{Today}} {{#Contacts}}{{FirstName}}{{/Contacts}}";
    expect(extractTags(text)).toEqual([
      "Account.Name",
      "Today",
      "#Contacts",
      "FirstName",
      "/Contacts",
    ]);
  });

  it("ignores single braces and empty tags", () => {
    expect(extractTags("{Account.Name} and {{}} stay out")).toEqual([]);
  });
});

describe("classifyTag", () => {
  it("resolves known keys, including format suffixes", () => {
    expect(classifyTag("Account.Name", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("Opportunity.Amount:currency", RESOLVABLE).status).toBe(
      "Resolved",
    );
  });

  it("classifies structural tags", () => {
    for (const inner of [
      "#Contacts",
      "/Contacts",
      "#if Opportunity.Amount > 1000",
      ":else",
      "/if",
      "^Account.HasDiscount__c",
    ]) {
      expect(classifyTag(inner, RESOLVABLE).status).toBe("Structural");
    }
  });

  it("flags unknown fields and suggests near-misses", () => {
    const wrongCase = classifyTag("account.name", RESOLVABLE);
    expect(wrongCase.status).toBe("Unresolved");
    expect(wrongCase.suggestion).toBe("Account.Name");

    const typo = classifyTag("Account.Industr", RESOLVABLE);
    expect(typo.status).toBe("Unresolved");
    expect(typo.suggestion).toBe("Account.Industry");
  });
});

describe("classifyDocumentTags (loop scope)", () => {
  it("resolves bare child fields inside a known loop", () => {
    const catalog = classifyDocumentTags(
      "{{#Contacts}}{{FirstName}} {{Email}}{{/Contacts}}",
      "Account",
    );
    expect(catalog.map((t) => t.status)).toEqual([
      "Structural",
      "Resolved",
      "Resolved",
      "Structural",
    ]);
  });

  it("flags typos inside a loop with an in-loop suggestion", () => {
    const catalog = classifyDocumentTags(
      "{{#Contacts}}{{FirstNme}}{{/Contacts}}",
      "Account",
    );
    expect(catalog[1]?.status).toBe("Unresolved");
    expect(catalog[1]?.suggestion).toBe("FirstName");
  });

  it("returns to root scope after the loop closes", () => {
    const catalog = classifyDocumentTags(
      "{{#Contacts}}{{FirstName}}{{/Contacts}} {{FirstName}}",
      "Account",
    );
    expect(catalog[3]?.status).toBe("Unresolved"); // bare key invalid at root
  });
});

describe("MockSliickClient.saveTemplate", () => {
  it("returns Valid when every tag resolves", async () => {
    const client = new MockSliickClient();
    client.documentText = "Hello {{Account.Name}} — {{#Contacts}}{{FirstName}}{{/Contacts}}";
    const result = await client.saveTemplate(
      { name: "T1", baseObjectApiName: "Account", fileName: "T1.docx" },
      "AAAA",
    );
    expect(result.validationStatus).toBe("Valid");
    expect(result.tagCatalog.map((t) => t.status)).toEqual([
      "Resolved",
      "Structural",
      "Resolved",
      "Structural",
    ]);
  });

  it("returns Invalid with a warning-free catalog when a tag is unknown", async () => {
    const client = new MockSliickClient();
    client.documentText = "Hello {{Account.Nmae}}";
    const result = await client.saveTemplate(
      { name: "T2", baseObjectApiName: "Account", fileName: "T2.docx" },
      "AAAA",
    );
    expect(result.validationStatus).toBe("Invalid");
    expect(result.tagCatalog[0]?.status).toBe("Unresolved");
  });

  it("warns when the document has no tags at all", async () => {
    const client = new MockSliickClient();
    client.documentText = "Plain document.";
    const result = await client.saveTemplate(
      { name: "T3", baseObjectApiName: "Account", fileName: "T3.docx" },
      "AAAA",
    );
    expect(result.warnings.some((w) => w.code === "NO_TAGS")).toBe(true);
  });

  it("upserts by name in listTemplates", async () => {
    const client = new MockSliickClient();
    client.documentText = "{{Account.Name}}";
    await client.saveTemplate(
      { name: "Same", baseObjectApiName: "Account", fileName: "a.docx" },
      "AAAA",
    );
    await client.saveTemplate(
      { name: "Same", baseObjectApiName: "Account", fileName: "b.docx" },
      "AAAA",
    );
    const list = await client.listTemplates("Account");
    expect(list.templates).toHaveLength(1);
    expect(list.templates[0]?.fileName).toBe("b.docx");
  });
});
