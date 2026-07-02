# Sliick Docs for Microsoft Word — User Guide

Build Salesforce document templates **inside Microsoft Word**. You write the
document the way you want it to look, drop in merge fields from a live picker
that knows your org's objects and fields, and save it to Salesforce. From then
on, Sliick Docs uses it to generate finished Word documents — one at a time
from a record, or in bulk — with each merge field replaced by real data.

This guide is for the person authoring templates. No coding required.

---

## Contents

1. [What you can build](#1-what-you-can-build)
2. [Opening the add-in](#2-opening-the-add-in)
3. [Connecting to Salesforce](#3-connecting-to-salesforce)
4. [The task pane at a glance](#4-the-task-pane-at-a-glance)
5. [Inserting a merge field](#5-inserting-a-merge-field)
6. [Inserting a repeating table (related lists)](#6-inserting-a-repeating-table-related-lists)
7. [Showing or hiding content with conditions](#7-showing-or-hiding-content-with-conditions)
8. [Saving to Salesforce](#8-saving-to-salesforce)
9. [Fixing tag problems](#9-fixing-tag-problems)
10. [Previewing with real data](#10-previewing-with-real-data)
11. [After you save: publishing and generating](#11-after-you-save-publishing-and-generating)
12. [Authoring tips for clean output](#12-authoring-tips-for-clean-output)
13. [What's supported (and what isn't yet)](#13-whats-supported-and-what-isnt-yet)
14. [Troubleshooting](#14-troubleshooting)
15. [Merge tag cheat sheet](#15-merge-tag-cheat-sheet)
16. [Appendix A — Admin setup: connect your org](#appendix-a--admin-setup-connect-your-org)

---

## 1. What you can build

A Sliick Docs template is an ordinary Word document with **merge fields** in it.
A merge field is a placeholder like `{{Account.Name}}` that gets replaced with
the real value when the document is generated.

You can insert four kinds of dynamic content:

- **Fields** — a single value from the record (account name, amount, close date)
  or a related parent (the account owner's email).
- **Repeating tables** — one table row per related record (every contact on an
  account, every line item on an opportunity).
- **Conditions** — content that only appears when a rule is true ("show the
  premium clause only when revenue is over $50,000").
- **Built-ins** — today's date, the current user's name, and similar.

Everything else in the document — your fonts, logos, tables, headers, footers,
page layout — is preserved exactly as you authored it.

---

## 2. Opening the add-in

How you get the add-in depends on how your organization distributes it:

- **Deployed by your admin** — open Word, go to the **Home** tab, and click
  **Merge Fields** in the Sliick Docs group on the ribbon. (If you don't see
  it, your admin may need to deploy it to your account first.)
- **Manual sideload (testing)** — Word → **Insert → Add-ins → My Add-ins →
  Upload My Add-in**, and choose the `manifest.xml` file you were given. On
  Word for the web it's **Add-ins → Upload My Add-in**.

Once installed, the **Merge Fields** button opens the Sliick Docs task pane
down the right-hand side of Word. You can work in the document and the pane at
the same time.

---

## 3. Connecting to Salesforce

The first time you open the pane it runs in **Demo mode** — it shows sample
Account, Contact, and Opportunity data so you can try everything without a
connection. The footer at the bottom of the pane tells you which mode you're in.

To work with your real org:

1. Click the **gear icon (⚙)** at the top of the pane.
2. Untick **Demo mode**.
3. Enter your **Salesforce org URL** (your My Domain address, e.g.
   `https://yourcompany.my.salesforce.com`).
4. Enter the **External Client App consumer key** — your Salesforce admin
   provides this from your org's External Client App (see
   [Appendix A](#appendix-a--admin-setup-connect-your-org)).
5. Click **Save**, then **Sign in to Salesforce**.

A Salesforce login window opens. Sign in as you normally would and approve
access. The pane then loads your org's objects and fields. You stay signed in
between sessions; use **Sign out** in Settings to disconnect.

> Everything the add-in shows you respects your Salesforce permissions. You only
> see objects and fields you already have access to.

> **Admin?** A one-time org setup is required before authors can connect:
> install the managed package, enable CORS for OAuth endpoints, and create your
> org's External Client App — full steps in
> [Appendix A](#appendix-a--admin-setup-connect-your-org).

---

## 4. The task pane at a glance

- **Object selector** (top dropdown) — the Salesforce object your template is
  built for. Pick this first; it drives every field below. A template for
  Accounts is generated from an Account record, a template for Opportunities
  from an Opportunity, and so on.
- **Search box** — type to filter the fields below by name.
- **Field groups** — collapsible sections:
  - **Fields** — the object's own fields.
  - **Parent fields** — fields reached through a lookup (e.g. the account's
    owner or parent account).
  - **Running user** — details of whoever generates the document.
  - **Built-ins** — today's date, current date/time.
  - **Related lists (repeating tables)** — child records you can loop over.
- **Action buttons** — **+ Condition**, **+ If blank**, **+ Image**,
  **+ Barcode**, **Review tags**, **My templates**, and **Save to Salesforce**.
- **✨ Insert sample layout** — drops a ready-made example (heading fields, a
  repeating table over the first related list, and totals) so you can see a
  working template in one click.
- **Smart cursor** — put your cursor inside a repeating list and the pane
  switches to that list's fields; select an existing tag and an **Edit**
  card appears that reopens it in the matching wizard.

---

## 5. Inserting a merge field

1. Click in your Word document where you want the value to appear.
2. In the pane, find the field (use the search box if it's a long list).
3. Click the field. The tag is inserted at your cursor — for example
   `{{Account.Name}}`.

That's it. When the document is generated, `{{Account.Name}}` becomes the actual
account name.

**Formatting is automatic.** The add-in adds the right format for the field
type so numbers and dates come out looking right:

- Currency fields → `{{Opportunity.Amount:currency}}` → **$50,000.00**
- Percent fields → `{{Opportunity.Probability:percent}}` → **75%**
- Date fields → `{{Opportunity.CloseDate:MM/dd/yyyy}}` → **03/18/2026**
- Checkbox fields → `{{Account.Active__c:checkbox}}` → **[X]** or **[ ]**

You can edit these tags by hand if you want a different date format — anything
Word can show as text, you can type as a tag.

**Parent fields** reach across relationships, up to five hops:
`{{Account.Owner.Email}}` pulls the account owner's email.

**Insert with options (⋯).** Every field row has a **⋯** button that opens an
options panel: pick a different **format**, choose a **locale** (e.g. German
formatting — `{{Amount:currency:de_DE}}` renders 1.234,50), set an **"if
blank, show"** fallback text (`{{Description|No notes on file}}`), and **📌
pin** the field to the top of the list. Recently used fields also float to a
**Recent** section automatically. In the search box, **Enter** inserts the top
match.

---

## 6. Inserting a repeating table (related lists)

Use this for line items, contacts, or any list of child records.

1. Put your cursor where the table should go.
2. In the pane, open **Related lists (repeating tables)** and click the
   relationship you want (e.g. **Contacts** or **Products (Line Items)**).
3. A wizard lists the child fields. Tick the columns you want.
4. Click **Insert table**.

The add-in inserts a two-row table: a **header row** with your column labels,
and a **data row** that repeats once per child record when the document is
generated. So three contacts produce three rows automatically.

You can restyle the table (borders, shading, widths) like any Word table — just
don't delete the merge fields in the data row.

**Filter and sort.** The wizard's optional filter section limits which records
appear (`{{#Opportunities WHERE Opportunities.Amount > 10000}}`) and sorts them
(`ORDER BY CloseDate DESC`) — pick a field, operator, and value, plus a sort
field and direction. To change a list's filter later, click its `{{#…}}` tag in
the document and use the **Edit** card.

**Nested lists (one level).** The same wizard can insert a list inside a list —
for example each opportunity, then each of its line items. Under **Nested
related list**, pick the child's own related list, tick the fields for each
level, and click **Insert nested list**. Each parent gets a heading line
followed by a **real table** of its child records (the classic grouped-invoice
layout); the outer `{{#…}}` / `{{/…}}` tags sit on their own paragraphs around
the table — keep them there.

> **Don't hand-type a nested list inside a table cell.** A `{{#…}}…{{/…}}` pair
> placed inside one cell of a row that already repeats would duplicate the whole
> row and blank its other fields, so Save now rejects that shape with a fix-it
> message. Put the outer repeat on paragraphs around the table instead — the
> shape the wizard inserts.

**Approval history.** Every object's related lists include **Approval
History** — a repeating table over the record's approval steps (actor, step
status, comments, dates). It works like any other related list; totals aren't
available for it.

---

## 7. Showing or hiding content with conditions

### Conditional content (if / else)

1. Optionally select the text you want to make conditional.
2. Click **+ Condition**.
3. Choose a **field**, an **operator** (=, ≠, >, <, ≥, ≤, or **contains** for
   "text appears anywhere in the field"), and a **value**.
4. Tick **NOT** to invert the test ("show when this is *false*"), and
   **Include an otherwise (else) branch** if you want alternate text when the
   condition is false. A second clause can be joined with **AND** / **OR**.
5. Click **Insert**.

Example: a condition on Annual Revenue `>` `50000` wraps your selected text so
it only appears for high-revenue accounts. With an else branch, you get
"Premium tier" when true and "Standard tier" when false.

### Show only when blank (if blank)

Click **+ If blank**, pick a field, and the content between the tags appears
**only when that field is empty** — handy for "No notes on file" style
fallbacks.

---

## 7b. Barcodes, and checking your tags

**Barcodes / QR codes.** Click **+ Barcode**, pick the field whose value gets
encoded (an invoice number, a URL), choose Code 128 or QR, and optionally a
size. The barcode image is generated into the document at merge time
(`{{*Account.AccountNumber}}`, `{{*Account.Website:qr:150}}`).

**Review tags.** Click **Review tags** any time for an instant check of every
tag in the document against your org — typos get "did you mean" suggestions,
fields used in the wrong list scope are flagged, and clicking any row jumps to
that tag in the document. A **Highlight tags** toggle shades every tag in the
document while you work (the shading is removed automatically when you save).
The full server-side validation still runs on save.

---

## 8. Saving to Salesforce

When your template is ready:

1. Click **Save to Salesforce**.
2. Give the template a **name** (defaults to the document's file name).
3. The **base object** is shown (set by your object selector — read-only).
4. Optionally enter a **Test record Id** — a real record used for Preview
   (see [section 10](#10-previewing-with-real-data)).
5. Click **Save template**.

**Updating an existing template.** Click **My templates** to see the templates
already saved for this base object. Pick one and the next save creates a **new
version** of it instead of a separate template (the save panel says whose
version you're saving, with a "save as a new template instead" escape hatch).
After any save, further saves from the same pane keep versioning that template.

The add-in uploads the document and checks every tag against your org. You then
see the **Save results** panel:

- A **green banner** means the template is valid and ready to publish.
- A **red banner** means some tags couldn't be matched — fix them and save again.
- Below the banner, **every tag is listed with its status**:
  - **Resolved** — matches a real field. ✓
  - **Structural** — a loop or condition tag (these are expected). ✓
  - **Unresolved** — no matching field; this needs fixing. ✗
  - **FLS warning** — the field exists but the running user may not see it.

For any unresolved tag, the panel suggests a likely correction
("Did you mean `Account.Industry`?").

---

## 9. Fixing tag problems

Unresolved tags almost always come from a typo or a field that doesn't exist on
the chosen object:

- Check the spelling against the suggestion in the Save results panel.
- Re-insert the field from the picker instead of typing it by hand — that
  guarantees a valid tag.
- Make sure you picked the right **base object** at the top. A field that's
  valid on Opportunity won't resolve on a template built for Account.

Fix the document in Word, click **Back to fields**, then **Save to Salesforce**
again. The check re-runs.

---

## 10. Previewing with real data

From the Save results panel, click **Preview** to generate the document against
real data, then download the finished `.docx` and open it in Word. This is the
fastest way to confirm fields, tables, and conditions all behave before anyone
else uses the template.

The **Preview record Id** box lets you preview against any record of the base
object — leave it as-is to use the test record you set at save time, or paste a
different record Id to try another one without re-saving.

> Preview needs a live Salesforce connection — it isn't available in Demo mode.

---

## 11. After you save: publishing and generating

Saving puts the template into Salesforce as a draft. From there it follows the
normal Sliick Docs lifecycle, handled inside Salesforce (not in Word):

1. **Publish** it from the **Sliick Docs → Template Library**. Publishing makes
   it available for generation. (An admin or template author usually does this.)
2. **Generate** finished documents:
   - **From a record** — the Generate button on an Account, Opportunity, etc.
   - **In bulk** — the Batch console, to produce one document per record across
     many records.
   - **From a Flow** — as part of an automated process.

Generated documents are saved to **Files** on the record and can be emailed
automatically, depending on how your admin has set things up.

**Output: Word (.docx) — and PDF when the template supports it.** Office
templates always generate Word, preserving your exact formatting. **PDF** is
also available for templates that are "PDF-ready" (see below). Which format is
produced is chosen **at generation time in Salesforce** — on the record's
Generate button, in a Flow, or in the Batch console — not in Word.

After you save, the results panel tells you whether the template is **PDF-ready**.
Some Word features don't survive native PDF rendering (text boxes, custom fonts,
floating images, SmartArt/shapes); if your template uses them, it stays
**Word-only** and the panel lists what to change — the Word output is unaffected
either way. Stick to tables, inline images, and standard fonts to keep a
template PDF-ready.

You can revise a template any time: open the document in Word, make changes,
and **Save to Salesforce** again to create a new version.

---

## 12. Authoring tips for clean output

A few Word habits keep generated documents looking right:

- **Use tables, not text boxes**, for side-by-side layout. Text boxes can shift
  position when documents are generated.
- **Anchor images "In Line with Text"** rather than floating them.
- **Set fixed column widths** on tables (Table Layout → AutoFit → Fixed Column
  Width). "AutoFit to Contents" can make columns jump around.
- **Accept or reject all tracked changes** before saving. Tracked edits can
  render as visible text. (If you save with tracked changes, the add-in warns
  you.)
- **Keep the file under 10 MB.** If it's bigger, compress images
  (Picture Format → Compress Pictures).
- **Place conditional text inline** within a sentence where you can, so that
  when it's hidden the surrounding text still flows naturally.

---

## 13. What's supported (and what isn't yet)

**Supported now:**

- Single fields, parent fields (up to 5 levels), and built-ins
- Automatic formatting for currency, percent, date, date-time, checkboxes, and
  **picklist labels** (picklist fields show the label, not the API value)
- Repeating tables over child relationships, including **one level of nested
  loops** (e.g. each opportunity, then each of its line items)
- **Totals** — SUM / COUNT / AVG / MIN / MAX over a related list (insert from
  the repeating-table panel; place the total outside the table)
- Conditions (if / else) with =, ≠, ordering, **contains**, and **NOT**;
  **compound conditions** (two clauses joined by AND/OR); "show when blank";
  and approval history loops
- **Images** from a field that holds a Salesforce File (the **+ Image** button),
  embedded into the document at generation time, with optional size
- Save-time validation with suggestions, preview against any record, and a
  **template library** for saving new versions of existing templates

**Not in this version:**

- **Excel and PowerPoint** — Word only for now.
- **More than one level of loop nesting** (a loop inside a loop inside a loop).
- **PDF for non-PDF-ready templates** — templates using text boxes, custom
  fonts, floating images, or SmartArt stay Word-only (the save panel says so).
- **Barcodes / QR codes.**

The pane only offers what your installed Sliick Docs version supports, so you
won't be shown options that won't work.

---

## 14. Troubleshooting

**The Merge Fields button isn't on the ribbon.**
The add-in isn't installed for your account. Ask your admin to deploy it, or
sideload the manifest (see [section 2](#2-opening-the-add-in)).

**Sign-in window opens then closes / I can't connect.**
Double-check the org URL and consumer key in Settings. The org URL must be your
My Domain HTTPS address. If it still fails, your admin may need to confirm the
External Client App and its allowed origins are set up.

**A field I expect isn't in the list.**
You may not have permission to see it in Salesforce, or you're on the wrong base
object. Switch the object selector or ask your admin about field access.

**Save says the template is invalid.**
One or more tags didn't match. Open the Save results panel, read the per-tag
status, apply the suggestions, and save again.

**Preview is greyed out.**
You're in Demo mode — preview needs a live Salesforce connection. Once
connected, save and use the Preview record Id box on the Save results panel.

**My table shows only one row in the output.**
Make sure the repeating fields sit inside a row inserted by the **Related lists**
wizard. A table typed by hand without the loop tags won't repeat.

---

## 15. Merge tag cheat sheet

You'll normally insert these from the pane, but here's what they look like if
you ever need to read or type one. String values in conditions use **single
quotes**.

| Purpose | Example |
|---|---|
| Field | `{{Account.Name}}` |
| Parent field | `{{Account.Owner.Email}}` |
| Currency / percent / date | `{{Opportunity.Amount:currency}}` · `{{Opportunity.Probability:percent}}` · `{{Opportunity.CloseDate:MM/dd/yyyy}}` |
| Checkbox | `{{Account.Active__c:checkbox}}` |
| Today / now | `{{Today}}` · `{{Now}}` |
| Current user | `{{RunningUser.Name}}` · `{{RunningUser.Email}}` |
| Picklist label | `{{Opportunity.StageName:label}}` |
| Repeating table | `{{#Contacts}}` … `{{FirstName}}` … `{{/Contacts}}` |
| Nested loop | `{{#Opportunities}}` … `{{#OpportunityLineItems}}` … `{{/OpportunityLineItems}}` … `{{/Opportunities}}` |
| Total | `{{SUM:OpportunityLineItems.TotalPrice:currency}}` · `{{COUNT:Contacts}}` |
| Condition | `{{#if Account.AnnualRevenue > 50000}}` … `{{:else}}` … `{{/if}}` |
| Compound condition | `{{#if Account.AnnualRevenue > 50000 AND Account.Industry = 'Technology'}}` … `{{/if}}` |
| Contains / NOT | `{{#if Account.Description contains 'priority'}}` · `{{#if NOT (Opportunity.StageName = 'Closed Won')}}` |
| Filtered / sorted list | `{{#Opportunities WHERE Opportunities.Amount > 10000 ORDER BY CloseDate DESC}}` … `{{/Opportunities}}` |
| Fallback text | `{{Account.Description\|No notes on file}}` · `{{Amount:currency\|N/A}}` |
| Locale format | `{{Amount:currency:de_DE}}` · `{{CloseDate:date:fr_FR}}` |
| Barcode / QR | `{{*Account.AccountNumber}}` · `{{*Account.Website:qr:150}}` |
| Approval history | `{{#Approvals}}` … `{{ActorName}}` `{{StepStatus}}` … `{{/Approvals}}` |
| Show when blank | `{{^Account.Description}}` … `{{/Account.Description}}` |
| Image | `{{%Account.Logo__c}}` · `{{%Account.Logo__c:200x60}}` |

Inside a repeating table, child fields are written **without** the object name —
`{{FirstName}}`, not `{{Contact.FirstName}}`. The same applies to fields and
totals inside a nested loop. Totals (`SUM`/`COUNT`/…) go **outside** the table,
and string values in conditions use **single quotes**.

---

## Appendix A — Admin setup: connect your org

This appendix is for **Salesforce admins**. Template authors don't need it —
they only enter the org URL (and, optionally, a consumer key) as in
[Section 3](#3-connecting-to-salesforce).

Connecting an org takes two steps. The first is required; the second is optional.

Each org connects with **its own External Client App** — you create it once, and
your authors enter its consumer key. There are three one-time steps: install the
package, enable CORS for the OAuth endpoints, and create the External Client App.

### A.1 Install the Sliick Docs managed package (required)

The add-in reads merge fields and saves templates through Sliick Docs' API,
which lives in the **Sliick Docs managed package**. Installing it also adds
`https://office.sliick.com` to your org's **CORS allowlist** so the add-in's
browser requests are accepted. Without the package installed, the add-in can't
load your fields.

### A.2 Enable CORS for OAuth endpoints (required)

The add-in signs in directly from the browser, so the org must allow the OAuth
endpoints to respond to it. In **Setup → CORS**, turn on **"Enable CORS for
OAuth endpoints."** (The package already put `office.sliick.com` on the CORS
allowlist; this setting lets the sign-in/token endpoints use it.) Without this,
sign-in fails to complete.

### A.3 Create your org's External Client App (required)

In **Setup**, open **External Client App Manager → New External Client App**
(older orgs: **App Manager → New Connected App**). Configure:

- **Basic information**
  - **Name:** `Sliick Office Add-in` (anything you like)
  - **Contact email:** your admin email
  - **Distribution state:** **Local**
- **OAuth settings → Enable OAuth**
  - **Callback URL:**
    ```
    https://office.sliick.com/auth-callback.html
    ```
  - **OAuth scopes** (add all three):
    - *Manage user data via APIs* (`api`)
    - *Perform requests at any time* (`refresh_token`, `offline_access`)
    - *Access the identity URL service* (`openid`)
  - **Require Proof Key for Code Exchange (PKCE):** **ON**
  - **Require secret for Web Server Flow:** **OFF**
  - **Require secret for Refresh Token Flow:** **OFF**

  > The "Require secret…" boxes are off because the add-in is a **public
  > client** — it runs in the browser and holds no secret. PKCE secures the flow.

- **Save.** Salesforce generates a **Consumer Key**. You'll hand this to authors.

### A.4 Set the OAuth policies

After saving, open the app's **OAuth Policies** (Edit Policies) and set:

- **IP Relaxation:** **Relax IP restrictions.** Also leave the **Refresh Token
  IP Allowlist** / IP-enforcement options **off** for this app.

  > Authors sign in from wherever they work (home, mobile, office), so the OAuth
  > token grant comes from an arbitrary IP. With "Enforce," Salesforce can't
  > challenge a token grant for device verification, so it hard-blocks it with
  > *"ip restricted by app developer."* Relaxing IP for **this app** lets authors
  > connect from any IP; security is still enforced by PKCE and the standard
  > Salesforce login. (This is a per-app setting on your own org's app — it does
  > not change your org-wide login IP policy.)

- **Permitted Users:** choose **Admin approved users are pre-authorized** (then
  assign a permission set / profile to the authors who should connect) or
  **All users may self-authorize**, per your org's policy.

### A.5 Give authors the consumer key

Copy the **Consumer Key** from your External Client App and share it with your
template authors. In the add-in they open **Settings (⚙)**, untick **Demo
mode**, enter the **org URL**, paste the **consumer key**, and **Sign in**.

> The consumer key is a public client identifier, not a secret — it's safe to
> share with your authors.
