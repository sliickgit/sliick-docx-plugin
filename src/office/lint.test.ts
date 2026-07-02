import { describe, expect, it } from "vitest";
import { DiscoverResponse } from "../api/types";
import { buildLintContext, lintDocumentText } from "./lint";

const ROOT: DiscoverResponse = {
  baseObjectApiName: "Account",
  baseObjectLabel: "Account",
  rootScalarMergeFields: [
    { key: "Account.Name", label: "Account Name", type: "string" },
    { key: "Account.AnnualRevenue", label: "Annual Revenue", type: "currency" },
  ],
  parentLookupMergeFields: [{ key: "Account.Owner.Email", label: "Owner Email", type: "email" }],
  runningUserMergeFields: [{ key: "RunningUser.Name", label: "User Name", type: "string" }],
  builtInMergeFields: [{ key: "Today", label: "Today", type: "date" }],
  childRelationships: [
    { relationshipName: "Contacts", childObjectApiName: "Contact", label: "Contacts" },
    { relationshipName: "Approvals", childObjectApiName: "ProcessInstanceStep", label: "Approval History" },
  ],
};

const CONTACT: DiscoverResponse = {
  baseObjectApiName: "Contact",
  baseObjectLabel: "Contact",
  rootScalarMergeFields: [
    { key: "Contact.LastName", label: "Last Name", type: "string" },
    { key: "Contact.Email", label: "Email", type: "email" },
  ],
  parentLookupMergeFields: [],
  runningUserMergeFields: [],
  builtInMergeFields: [],
  childRelationships: [],
};

const ctx = () => buildLintContext(ROOT, new Map([["Contact", CONTACT]]));

describe("lintDocumentText", () => {
  it("resolves root fields, loop fields, and grammar-v2 forms", () => {
    const out = lintDocumentText(
      "{{Account.Name}} {{Account.AnnualRevenue:currency:de_DE|N/A}} {{Today}}" +
        " {{#Contacts WHERE Contacts.Email contains '@' ORDER BY LastName}}{{LastName}}{{/Contacts}}" +
        " {{*Account.Name:qr}} {{COUNT:Contacts}}",
      ctx(),
    );
    expect(out.errors).toEqual([]);
    expect(out.hasProblems).toBe(false);
    const statuses = new Map(out.entries.map((e) => [e.tag, e.status]));
    expect(statuses.get("{{Account.Name}}")).toBe("Resolved");
    expect(statuses.get("{{Account.AnnualRevenue:currency:de_DE|N/A}}")).toBe("Resolved");
    expect(statuses.get("{{LastName}}")).toBe("Resolved");
    expect(statuses.get("{{*Account.Name:qr}}")).toBe("Resolved");
    expect(statuses.get("{{COUNT:Contacts}}")).toBe("Resolved");
  });

  it("flags typos with suggestions, scoped to where the tag sits", () => {
    const out = lintDocumentText(
      "{{Account.Nmae}} {{#Contacts}}{{LastNme}}{{/Contacts}}",
      ctx(),
    );
    expect(out.hasProblems).toBe(true);
    expect(out.entries[0]).toMatchObject({ status: "Unresolved", suggestion: "Account.Name" });
    expect(out.entries[2]).toMatchObject({ status: "Unresolved", suggestion: "LastName" });
  });

  it("catches root-scope fields used inside a loop", () => {
    const out = lintDocumentText("{{#Contacts}}{{Account.Name}}{{/Contacts}}", ctx());
    expect(out.entries[1]!.status).toBe("Unresolved");
    expect(out.entries[1]!.note).toContain("Contacts");
  });

  it("validates loop modifiers: filter prefix and sort field", () => {
    const badPrefix = lintDocumentText(
      "{{#Contacts WHERE Email contains '@'}}{{LastName}}{{/Contacts}}",
      ctx(),
    );
    expect(badPrefix.errors.some((e) => e.includes("prefix"))).toBe(true);

    const badSort = lintDocumentText(
      "{{#Contacts ORDER BY Nope}}{{LastName}}{{/Contacts}}",
      ctx(),
    );
    expect(badSort.errors.some((e) => e.includes("Sort field"))).toBe(true);
  });

  it("flags unknown related lists, unbalanced blocks, and misplaced totals", () => {
    const unknown = lintDocumentText("{{#Contracts}}{{X}}{{/Contracts}}", ctx());
    expect(unknown.entries[0]).toMatchObject({ status: "Unresolved", suggestion: "Contacts" });

    const unbalanced = lintDocumentText("{{#Contacts}}{{LastName}}", ctx());
    expect(unbalanced.errors.some((e) => e.includes("never closed"))).toBe(true);

    const inLoopTotal = lintDocumentText(
      "{{#Contacts}}{{COUNT:Contacts}}{{/Contacts}}",
      ctx(),
    );
    expect(inLoopTotal.errors.some((e) => e.includes("outside"))).toBe(true);
  });

  it("resolves Approvals via the fixed synthetic field set", () => {
    const out = lintDocumentText(
      "{{#Approvals}}{{ActorName}} {{ActorNme}}{{/Approvals}}",
      ctx(),
    );
    expect(out.entries[1]!.status).toBe("Resolved");
    expect(out.entries[2]).toMatchObject({ status: "Unresolved", suggestion: "ActorName" });
  });

  it("is lenient for loops whose child discover is not loaded yet", () => {
    const bare = buildLintContext(ROOT, new Map());
    const out = lintDocumentText("{{#Contacts}}{{Anything}}{{/Contacts}}", bare);
    expect(out.entries[1]!.status).toBe("Resolved");
  });
});
