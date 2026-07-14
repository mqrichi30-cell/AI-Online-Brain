# AWG – Ship with POs Report · Automation Setup

This automation replaces the manual "run the `ShipWith Report.xlsm` macros" step.
When an email with subject **`AWG - Ship with POs Repot`** arrives, Power Automate
parses the attached **Consolidations** report, matches each Ship-to name to a
contact, and **creates one draft per recipient** in the `pgcustservw2.im@pg.com`
mailbox — exactly what the old `CreateConsolidationEmails` macro did, but
unattended. Nothing is sent automatically; you review and send the drafts.

> Mode chosen: **Create drafts** (safe). To switch to auto-send later, see
> [Switching to auto-send](#switching-to-auto-send).

---

## How it works

```
Email arrives  (subject = "AWG - Ship with POs Repot", folder "AWG + Wakefern")
   │
   ├─ 1. Filter the .xlsx attachment (the Consolidations report)
   ├─ 2. Overwrite a FIXED temp .xlsx in OneDrive with the attachment bytes
   ├─ 3. List contacts from "Regional Team – Contacts Data Base.xlsx" (SharePoint)
   ├─ 4. Run Office Script `shipWithReport` on the temp file
   │        → parses consolidations, matches Ship-to → email,
   │          returns [{ to, cc, subject, htmlBody, matched }]
   └─ 5. For each item → POST to Graph /messages → DRAFT in shared mailbox
```

Why an Office Script? The block-parsing + HTML building is far too complex for
Power Automate expressions. The script (TypeScript) owns all the logic; the flow
just moves data and creates drafts. See
[`office-scripts/shipWithReport.ts`](../office-scripts/shipWithReport.ts).

The parsing/HTML logic is mirrored 1:1 in the runnable Python prototype
([`prototype/parse_and_preview.py`](../prototype/parse_and_preview.py)), which was
validated against the real attachment (10 consolidations → grouped drafts). Use it
to preview output without touching Power Automate.

---

## One-time setup

### A. Create the Office Script

1. Open **any** workbook in Excel on the web (excel.office.com) with the
   `pgcustservw2.im@pg.com` account (or your account that has access).
2. **Automate → New Script**.
3. Paste the entire contents of `office-scripts/shipWithReport.ts`. Save it as
   **`shipWithReport`**.
4. Note where it is stored (OneDrive). You'll pick it in the flow's *Run script*
   action.

### B. Create the fixed temp report file

The *Run script* action must point at a file that exists at design time, so we use
one fixed file that the flow overwrites every run.

1. In the `pgcustservw2.im@pg.com` OneDrive, create a folder e.g. `/AWG/`.
2. Put an empty Excel workbook there named **`ShipWith_Temp.xlsx`**.
   (Any valid .xlsx; its content is overwritten each run.)

### C. Confirm the Contacts Data Base is a Table

The flow reads contacts with *List rows present in a table*, which requires a
**formatted Excel Table**.

1. Open **Regional Team – Contacts Data Base.xlsx**
   (`.../NACSO-RegionalVMI/Shared Documents/General/Customer Documents/`).
2. Ensure the contact rows are inside a named Table (Insert → Table). Note the
   **table name** (e.g. `MasterData`).
3. Confirm the column headers for the **name** and **email**. The flow maps:
   - `name`  ← the **Ship to Name** column
   - `email` ← the **Email to** column

   If your headers differ, update the `Select_contacts` action mapping (see the
   flow) — the Office Script matches on `name` regardless.

> **Matching rule:** `Ship-to name1` in the report is compared to `Ship to Name`
> in the contacts DB after normalizing whitespace/case (so `AWG  GULF COAST`
> with a double space still matches `AWG GULF COAST`). A Ship-to with no contact
> gets its **own draft with a blank To** so you can address it by hand.

---

## Importing the flow package

A ready-to-import package is at
[`power-automate/AWG-ShipWithPOs-Repot.zip`](../power-automate/AWG-ShipWithPOs-Repot.zip).

1. **Power Automate → My flows → Import → Import Package (Legacy)**.
2. Upload the zip.
3. For each resource, pick the connection:
   - **Office 365 Outlook** → the existing `pgcustservw2.im@pg.com` connection.
   - **Excel Online (Business)** → create/select a connection.
   - **OneDrive for Business** → create/select a connection.
4. Import, then **open the flow to finish binding** (the package uses placeholders
   that must be replaced in the designer):

| Action | Field | Set to |
|---|---|---|
| Update temp report file (OneDrive) | File | `/AWG/ShipWith_Temp.xlsx` |
| List contacts (Excel) | Location / Document Library / File | the SharePoint **Contacts Data Base** workbook |
| List contacts (Excel) | Table | your contacts table name (e.g. `MasterData`) |
| Run script (Excel) | Location / File | `/AWG/ShipWith_Temp.xlsx` |
| Run script (Excel) | Script | `shipWithReport` |

The `<<PLACEHOLDER>>` values in `definition.json` (drive IDs, file IDs, script ID,
connection names) are resolved automatically when you re-pick those files/script
in the designer — you don't edit the JSON by hand.

> **Prefer building by hand?** The table above plus the diagram is enough to
> rebuild the flow from the existing skeleton. The import just saves you the
> clicks. If the legacy import is unavailable in your tenant, build it manually —
> the actions and expressions are all in `definition.json`.

---

## Draft creation detail

Drafts are created with the Office 365 Outlook **Send an HTTP request** action:

```
POST https://graph.microsoft.com/v1.0/users/pgcustservw2.im@pg.com/messages
Body: { subject, body{contentType:HTML, content}, toRecipients[], ccRecipients[] }
```

A `POST` to `/messages` always lands in **Drafts** (nothing is sent). The
connection identity needs **Send As / Full Access** on the shared mailbox — the
same permission the mailbox already uses.

**CC** is fixed to `marquardt.jw@pg.com; jackson.vs@pg.com` (from the original
macro). Change it in the `Compose_message` action if needed.

**Signature:** set the `SignatureHtml` variable at the top of the flow to your
real Outlook signature HTML (the old macro read a local `.htm` file, which a cloud
flow can't access).

---

## Switching to auto-send

To send instead of draft, replace the `Create_draft_in_shared_mailbox` action with
the Office 365 Outlook **Send an email from a shared mailbox (V2)** action:

- **Mailbox**: `pgcustservw2.im@pg.com`
- **To**: `@items('For_each_draft')?['to']`
- **Subject**: `@items('For_each_draft')?['subject']`
- **Body**: `@items('For_each_draft')?['htmlBody']` (Is HTML = Yes)
- **CC**: `marquardt.jw@pg.com; jackson.vs@pg.com`

Recommended: keep an `@items('For_each_draft')?['matched']` **condition** so only
matched recipients are auto-sent; unmatched ones stay as drafts for a human.

---

## Testing

1. Run the prototype locally to preview the exact emails:
   ```
   python3 prototype/parse_and_preview.py <report.xlsx> <contacts.xlsx>
   ```
   Open `samples/preview_all.html`.
2. In Power Automate, use **Test → Manually**, then send yourself a mail with the
   exact subject and the report attached.
3. Check the shared mailbox **Drafts** folder.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| No drafts created | Trigger folder/subject mismatch, or attachment isn't `.xlsx`. Subject must be exactly `AWG - Ship with POs Repot`. |
| Everyone in one draft with blank To | Contacts not loaded / table name wrong / column names don't match `Ship to Name` & `Email to`. |
| A known Ship-to is unmatched | Name differs beyond whitespace/case (e.g. `DIV` vs `DIVISION`). Align the name in the Contacts DB or add an alias row. |
| Run script fails "file not found" | The *Run script* / *Update file* actions point at different files. Both must be `/AWG/ShipWith_Temp.xlsx`. |
| Dates look like numbers | The script converts Excel serial dates; if a date column is text, it's passed through as-is. |
