# Add-missing-contact + auto-send

When a Ship-to has **no contact** in the Contacts Data Base, the main flow already
emails `pgcustservw2` with the full consolidation and a "Please add contact" note.
This feature lets you **add the contact once** and have it (1) saved to the
Contacts Data Base and (2) the email sent to the correct contact — without waiting
for the next report and without duplicating the ones that already had a contact.

## Design

```
Main flow (per unmatched draft):
   • sends notice to pgcustservw2 (as today), AND
   • creates an item in the "AWG ShipWith Pending" SharePoint list
     { Title = Ship-to, Subject, HtmlBody, Status = Pending }

You open the pending item and type the ContactEmail.

2nd flow  (trigger: item created/modified, ContactEmail filled, Status <> Sent):
   1. Add a row to the Contacts Data Base   (Ship to Name = Title, Email to = ContactEmail)
   2. Send the stored HtmlBody to ContactEmail (create message + send)
   3. Set the item Status = Sent
```

Because the contact is written to the Data Base, the next report matches it
automatically — you only ever fix a given customer once.

---

## Step 1 — Create the SharePoint list (you)

Site **NACSO-RegionalVMI** → **New → List** → name it **`AWG ShipWith Pending`**:

| Column | Type | Notes |
|---|---|---|
| Title | (exists) | holds the Ship-to name |
| ContactEmail | Single line of text | you type the email here |
| Subject | Single line of text | email subject |
| HtmlBody | Multiple lines of text (Plain text) | the composed email HTML |
| Status | Choice: `Pending`, `Sent` | default `Pending` |

---

## Step 2 — Main flow: create a pending item for unmatched drafts

Open **AWG - Ship with POs Repot** → inside **For each draft**, before/after
*Compose message*, add a **Condition**:

- **If** `items('For_each_draft')?['matched']` **is equal to** `false`
  (expression: `@equals(items('For_each_draft')?['matched'], false)`)

In the **If yes** branch add **SharePoint → Create item**:
- **Site Address**: NACSO-RegionalVMI
- **List Name**: AWG ShipWith Pending
- **Title**: `@{replace(items('For_each_draft')?['subject'], 'SHIP WITH NEEDED - ', '')}`
- **Subject**: `@{items('For_each_draft')?['subject']}`
- **HtmlBody**: `@{items('For_each_draft')?['htmlBody']}`
- **Status Value**: `Pending`

(The existing send actions stay as-is; unmatched still also go to pgcustservw2.)

---

## Step 3 — Build the 2nd flow "AWG ShipWith – Add Contact & Send"

**Trigger**: SharePoint **When an item is created or modified**
- Site: NACSO-RegionalVMI · List: AWG ShipWith Pending

**Condition** (guard):
- `@and(not(empty(triggerOutputs()?['body/ContactEmail'])), not(equals(triggerOutputs()?['body/Status/Value'], 'Sent')))`

**If yes** → add these actions in order:

1. **Excel Online (Business) → Add a row into a table**
   - Location / Library / File / Table = the **Contacts Data Base** (same as the
     main flow's *List contacts*)
   - **Ship to Name** = `@{triggerOutputs()?['body/Title']}`
   - **Email to** = `@{triggerOutputs()?['body/ContactEmail']}`
   - (leave *Ship to #* empty — matching is by name)

2. **Compose – message** (Compose):
   ```
   {
     "subject": "@{triggerOutputs()?['body/Subject']}",
     "importance": "normal",
     "body": { "contentType": "HTML", "content": "@{triggerOutputs()?['body/HtmlBody']}" },
     "toRecipients": "@json(concat('[{\"emailAddress\":{\"address\":\"', replace(replace(replace(trim(triggerOutputs()?['body/ContactEmail']), ' ', ''), ',', ';'), ';', '\"}},{\"emailAddress\":{\"address\":\"'), '\"}}]'))",
     "ccRecipients": "@json('[{\"emailAddress\":{\"address\":\"marquardt.jw@pg.com\"}},{\"emailAddress\":{\"address\":\"jackson.vs@pg.com\"}}]')"
   }
   ```

3. **Office 365 Outlook → Send an HTTP request** (create the message):
   - Method `POST` · Uri `v1.0/me/messages` · Body `@outputs('Compose_-_message')`
     (use the actual Compose action name)

4. **Office 365 Outlook → Send an HTTP request** (send it):
   - Method `POST`
   - Uri `v1.0/me/messages/@{encodeUriComponent(body('<create-step-name>')?['id'])}/send`
   - Body `{}`

5. **SharePoint → Update item**:
   - Site / List = AWG ShipWith Pending · Id = `@{triggerOutputs()?['body/ID']}`
   - Title = `@{triggerOutputs()?['body/Title']}` · **Status** = `Sent`

Use the connections that already exist: Office 365 Outlook (`pgcustservw2`),
SharePoint (`marin.c`), Excel Online Business (`marin.c`).
