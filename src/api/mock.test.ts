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

  it("classifies aggregate and image tags as resolvable", () => {
    expect(classifyTag("SUM:Opportunities.Amount", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("COUNT:Contacts", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("%Account.Name", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("%Account.Name:200x60", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("%Account.Nope", RESOLVABLE).status).toBe("Unresolved");
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

describe("classifyTag (grammar-v2 forms)", () => {
  it("resolves fallback pipes, barcodes, and locale suffixes like the engine", () => {
    expect(classifyTag("Account.Name|N/A", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("Account.Nmae|N/A", RESOLVABLE).status).toBe("Unresolved");
    expect(classifyTag("*Account.Name", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("*Account.Name:qr:150", RESOLVABLE).status).toBe("Resolved");
    expect(classifyTag("Opportunity.Amount:currency:de_DE", RESOLVABLE).status).toBe("Resolved");
  });

  it("treats loop opens with WHERE/ORDER BY as structural", () => {
    expect(
      classifyTag("#Contacts WHERE Contacts.Title = 'VP' ORDER BY LastName DESC", RESOLVABLE)
        .status,
    ).toBe("Structural");
  });
});

describe("classifyDocumentTags (grammar-v2 loop scope)", () => {
  it("scopes bare child fields inside a filtered loop", () => {
    const catalog = classifyDocumentTags(
      "{{#Contacts WHERE Contacts.Title = 'VP'}}{{FirstName}}{{/Contacts}}",
      "Account",
    );
    expect(catalog.map((t) => t.status)).toEqual(["Structural", "Resolved", "Structural"]);
  });

  it("keeps loop scope across an inverse block's close tag", () => {
    const catalog = classifyDocumentTags(
      "{{#Contacts}}{{^FirstName}}none{{/FirstName}}{{LastName}}{{/Contacts}}",
      "Account",
    );
    // {{LastName}} must still resolve in Contacts scope after {{/FirstName}}.
    expect(catalog[3]).toMatchObject({ tag: "{{LastName}}", status: "Resolved" });
  });
});

describe("capabilities (grammar-v2 parity)", () => {
  it("advertises the new backend flags", async () => {
    const caps = await new MockSliickClient().getCapabilities();
    expect(caps.features.loopFilters).toBe(true);
    expect(caps.features.fallbackText).toBe(true);
    expect(caps.features.localeFormats).toBe(true);
    expect(caps.features.barcodes).toBe(true);
    expect(caps.features.signatureTags).toBe(false);
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

  it("resolves the synthetic Approvals fields inside an Approvals loop", () => {
    const catalog = classifyDocumentTags(
      "{{#Approvals}}{{ActorName}} {{StepStatus}} {{ActedAt:MM/dd/yyyy}} {{ActorNme}}{{/Approvals}}",
      "Account",
    );
    expect(catalog.map((t) => t.status)).toEqual([
      "Structural",
      "Resolved",
      "Resolved",
      "Resolved",
      "Unresolved",
      "Structural",
    ]);
    expect(catalog[4]?.suggestion).toBe("ActorName");
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
    // office-pdf: save response carries the native-PDF verdict.
    expect(result.pdfReady).toBe(true);
    expect(result.pdfWarnings).toEqual([]);
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

  it("revises by templateId: keeps the id, bumps the version, allows rename", async () => {
    const client = new MockSliickClient();
    client.documentText = "{{Account.Name}}";
    const first = await client.saveTemplate(
      { name: "Quote", baseObjectApiName: "Account", fileName: "q.docx" },
      "AAAA",
    );
    const second = await client.saveTemplate(
      {
        templateId: first.templateId,
        name: "Quote v2",
        baseObjectApiName: "Account",
        fileName: "q2.docx",
      },
      "AAAA",
    );
    expect(second.templateId).toBe(first.templateId);
    expect(second.versionId).not.toBe(first.versionId);
    const list = await client.listTemplates("Account");
    expect(list.templates).toHaveLength(1);
    expect(list.templates[0]?.name).toBe("Quote v2");
    expect(list.templates[0]?.latestVersionId).toBe(second.versionId);
  });

  it("revising a template does not touch others with the same name history", async () => {
    const client = new MockSliickClient();
    client.documentText = "{{Account.Name}}";
    const a = await client.saveTemplate(
      { name: "A", baseObjectApiName: "Account", fileName: "a.docx" },
      "AAAA",
    );
    await client.saveTemplate(
      { name: "B", baseObjectApiName: "Account", fileName: "b.docx" },
      "AAAA",
    );
    await client.saveTemplate(
      { templateId: a.templateId, name: "A", baseObjectApiName: "Account", fileName: "a2.docx" },
      "AAAA",
    );
    const list = await client.listTemplates("Account");
    expect(list.templates).toHaveLength(2);
    expect(list.templates.find((t) => t.templateId === a.templateId)?.fileName).toBe("a2.docx");
    expect(list.templates.find((t) => t.name === "B")?.fileName).toBe("b.docx");
  });

  it("rejects a revise for an unknown templateId", async () => {
    const client = new MockSliickClient();
    client.documentText = "{{Account.Name}}";
    await expect(
      client.saveTemplate(
        { templateId: "nope", name: "X", baseObjectApiName: "Account", fileName: "x.docx" },
        "AAAA",
      ),
    ).rejects.toThrow(/no template/i);
  });
});
