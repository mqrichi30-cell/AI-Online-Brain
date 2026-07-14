# AI-Online-Brain

Automations for the NACSO / AWG + Wakefern team.

## AWG – Ship with POs Report

When an email with subject **`AWG - Ship with POs Repot`** arrives, this
automation parses the attached **Consolidations** report, matches each Ship-to
name to a contact, and **creates one draft email per recipient** in the
`pgcustservw2.im@pg.com` shared mailbox — replacing the manual
`ShipWith Report.xlsm` macro run. Nothing is sent; you review and send the drafts.

### Layout

| Path | What it is |
|---|---|
| `office-scripts/shipWithReport.ts` | Office Script (Excel Online) — parses the report, matches contacts, builds the draft HTML. All the logic lives here. |
| `power-automate/AWG-ShipWithPOs-Repot.zip` | Importable Power Automate package (updated flow: trigger → save attachment → run script → create drafts). |
| `power-automate/AWG-ShipWithPOs/` | Unpacked flow package (edit `definition.json` here, then re-zip). |
| `prototype/parse_and_preview.py` | Runnable reference of the same logic. Validates parsing and renders email previews without Power Automate. |
| `samples/` | Sample drafts + rendered `preview_all.html` generated from the real attachment. |
| `docs/SETUP.md` | Full setup / import / troubleshooting guide. |

### Quick preview (no Power Automate needed)

```bash
pip install openpyxl
python3 prototype/parse_and_preview.py path/to/report.xlsx [path/to/contacts.xlsx]
# then open samples/preview_all.html
```

### Setup

See **[docs/SETUP.md](docs/SETUP.md)**.

### Flow at a glance

```
Email arrives (subject "AWG - Ship with POs Repot")
  → filter .xlsx attachment
  → overwrite fixed temp .xlsx in OneDrive
  → list contacts from "Regional Team – Contacts Data Base.xlsx"
  → Run Office Script shipWithReport  →  [{ to, cc, subject, htmlBody }]
  → for each → create DRAFT in shared mailbox (Graph POST /messages)
```
