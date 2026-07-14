#!/usr/bin/env python3
"""
Prototype / reference implementation of the "AWG - Ship with POs Report" automation.

Purpose
-------
This script replicates, in a runnable form, the exact logic that will live inside
the production Office Script (office-scripts/shipWithReport.ts). It exists so the
parsing + grouping + email-HTML logic can be validated against the REAL email
attachment before it is deployed into Power Automate / Excel Online.

Input
-----
The "Consolidations" sheet of the report that arrives attached to the email whose
subject is "AWG - Ship with POs Repot". That report already contains the
consolidations (grouped, with a subtotal row after each block and a blank
separator row). Every row in it is a PO that needs a "Ship With".

What it produces
----------------
1) samples/draft_emails.json  -> the per-recipient draft payloads
                                 [{ to, cc, subject, htmlBody }, ...]
2) samples/preview_<n>.html   -> a rendered preview of each draft email body

Grouping rules (faithful to the original VBA `CreateConsolidationEmails`)
------------------------------------------------------------------------
- Each block in the report (consecutive data rows ended by a subtotal row) is
  ONE consolidation and gets a sequential Consolidation Number.
- Consolidations are grouped into emails by the resolved recipient email
  address (looked up from the Contacts Data Base by Ship-to name). One draft per
  recipient, containing every consolidation that belongs to that recipient.
- Subject: "SHIP WITH NEEDED - " + the distinct Ship-to names joined by "/".
- CC: the fixed distribution list.
"""

import json
import os
import sys
from html import escape

import openpyxl

# --- Configuration that mirrors the VBA constants -------------------------------

CC_LIST = "marquardt.jw@pg.com; jackson.vs@pg.com"

INTRO_HTML = (
    "Dear Team,<br><br>"
    "The following PO is undersized and requires a Ship With to make a full "
    "truckload. Do you prefer to write a Ship With PO we can consolidate OR do "
    "you prefer we add or increase product already on the PO?<br><br>"
)

# Placeholder signature. In production the flow injects the real Outlook signature
# (see docs/SETUP.md). Kept here so the preview looks complete.
SIGNATURE_HTML = (
    "<br>Best regards,<br>"
    "<strong>AWG + Wakefern Team</strong><br>"
    "P&amp;G Customer Service<br>"
)

# Column layout of the attachment "Consolidations" sheet (0-based).
COL_DB = 0            # DB / Delivery block
COL_CREATED_ON = 1    # Created_On
COL_GI = 2            # Goods Issue Date
COL_RDD = 3           # Requested Delivery Date
COL_PLANT = 4         # Plant
COL_MIX = 5           # Material Mix
COL_SHIPTO = 6        # Ship-to name1
COL_DOC = 7           # Sales document  -> Document Number
COL_PO = 8            # PO              -> Purchase Order Number
COL_SHIPCOND = 9      # Shipping Conditions
COL_WEIGHT = 10       # Gross Weight
COL_COF = 11          # COF-Cube Order Factor
COL_FP = 12           # Floor Position  -> "FP" in the email
COL_TM = 13           # Transport Medium
COL_STATUS = 14       # Status          -> shown as "Description"


def norm_name(value):
    """Normalize a Ship-to name for matching (collapse repeated whitespace, trim,
    uppercase). Handles the double-spaces seen in the source data e.g.
    'AWG  GULF COAST'."""
    if value is None:
        return ""
    return " ".join(str(value).split()).upper()


def fmt_date(value):
    """Render a date cell as M/D/YYYY (US style, matching the workbook)."""
    if value is None:
        return ""
    try:
        return f"{value.month}/{value.day}/{value.year}"
    except AttributeError:
        return str(value)


def fmt_num(value, decimals=3):
    if value is None or value == "":
        return ""
    try:
        return f"{float(value):,.{decimals}f}"
    except (ValueError, TypeError):
        return str(value)


def is_data_row(row):
    """A data row has both a DB value and a Ship-to name."""
    return bool(str(row[COL_DB]).strip()) and bool(str(row[COL_SHIPTO] or "").strip())


def is_subtotal_row(row):
    """A subtotal row has no DB/Ship-to but does carry a Gross Weight."""
    return (
        not str(row[COL_DB] or "").strip()
        and not str(row[COL_SHIPTO] or "").strip()
        and row[COL_WEIGHT] not in (None, "")
    )


def parse_consolidations(ws):
    """Parse the Consolidations sheet into a list of consolidation blocks.
    Each block is a dict: { number, ship_to, rows: [ {..cells..} ] }."""
    blocks = []
    current = []
    consolidation_number = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        # Pad the row so fixed indexes never go out of range.
        row = list(row) + [None] * (COL_STATUS + 1 - len(row))

        if is_data_row(row):
            current.append(row)
        else:
            # subtotal or blank -> close the current block if it has data
            if current:
                consolidation_number += 1
                blocks.append(_build_block(consolidation_number, current))
                current = []

    if current:
        consolidation_number += 1
        blocks.append(_build_block(consolidation_number, current))

    return blocks


def _build_block(number, rows):
    return {
        "number": number,
        "ship_to": rows[0][COL_SHIPTO],
        "rows": [
            {
                "created_on": r[COL_CREATED_ON],
                "gi": r[COL_GI],
                "rdd": r[COL_RDD],
                "plant": r[COL_PLANT],
                "mix": r[COL_MIX],
                "ship_to": r[COL_SHIPTO],
                "doc": r[COL_DOC],
                "po": r[COL_PO],
                "ship_cond": r[COL_SHIPCOND],
                "weight": r[COL_WEIGHT],
                "fp": r[COL_FP],
                "status": r[COL_STATUS],
            }
            for r in rows
        ],
    }


def build_block_table_html(block):
    """Render one consolidation as an HTML table (mirrors the VBA layout)."""
    head = (
        "<tr>"
        "<th style='background-color:yellow;'>Created On</th>"
        "<th style='background-color:yellow;'>Goods Issue Date</th>"
        "<th style='background-color:yellow;'>Request Delivery Day</th>"
        "<th style='background-color:yellow;'>Plant</th>"
        "<th>Material Mix</th>"
        "<th>Ship to Name</th>"
        "<th>Document Number</th>"
        "<th>Purchase Order Number</th>"
        "<th>Shipping Condition</th>"
        "<th style='background-color:yellow;'>Gross Weight</th>"
        "<th style='background-color:yellow;'>FP</th>"
        "<th>Description</th>"
        "</tr>"
    )

    body = []
    total_weight = 0.0
    total_fp = 0.0
    for r in block["rows"]:
        try:
            total_weight += float(r["weight"] or 0)
        except (ValueError, TypeError):
            pass
        try:
            total_fp += float(r["fp"] or 0)
        except (ValueError, TypeError):
            pass
        body.append(
            "<tr>"
            f"<td>{escape(fmt_date(r['created_on']))}</td>"
            f"<td>{escape(fmt_date(r['gi']))}</td>"
            f"<td>{escape(fmt_date(r['rdd']))}</td>"
            f"<td>{escape(str(r['plant'] or ''))}</td>"
            f"<td>{escape(str(r['mix'] or ''))}</td>"
            f"<td>{escape(str(r['ship_to'] or ''))}</td>"
            f"<td>{escape(str(r['doc'] or ''))}</td>"
            f"<td>{escape(str(r['po'] or ''))}</td>"
            f"<td>{escape(str(r['ship_cond'] or ''))}</td>"
            f"<td>{escape(fmt_num(r['weight']))}</td>"
            f"<td>{escape(fmt_num(r['fp']))}</td>"
            f"<td>{escape(str(r['status'] or ''))}</td>"
            "</tr>"
        )

    totals = (
        "<tr>"
        "<td colspan='9'><strong>Totals:</strong></td>"
        f"<td><strong>{total_weight:,.2f}</strong></td>"
        f"<td><strong>{total_fp:,.2f}</strong></td>"
        "<td></td>"
        "</tr>"
    )

    return (
        f"<h3>Consolidation Number: {block['number']}</h3>"
        "<table border='1' cellpadding='5' cellspacing='0'>"
        + head + "".join(body) + totals +
        "</table><br>"
    )


def build_drafts(blocks, contacts):
    """Group consolidation blocks by resolved recipient and build a draft payload
    per recipient. `contacts` maps normalized Ship-to name -> email address."""
    # recipient email -> { ship_to_names: set, blocks: [] }
    grouped = {}
    unmatched = []

    for block in blocks:
        key = norm_name(block["ship_to"])
        email = contacts.get(key, "")
        if email:
            group_key = email
        else:
            # Unmatched Ship-to: give each its own draft (blank recipient) so a
            # human can address it individually, instead of lumping them together.
            unmatched.append(block["ship_to"])
            group_key = "__UNMATCHED__" + key
        grouped.setdefault(group_key, {"email": email, "names": [], "blocks": []})
        if block["ship_to"] not in grouped[group_key]["names"]:
            grouped[group_key]["names"].append(block["ship_to"])
        grouped[group_key]["blocks"].append(block)

    drafts = []
    for data in grouped.values():
        body = INTRO_HTML
        for block in data["blocks"]:
            body += build_block_table_html(block)
        body += SIGNATURE_HTML
        subject = "SHIP WITH NEEDED - " + "/".join(data["names"])
        drafts.append(
            {
                "to": data["email"],
                "cc": CC_LIST,
                "subject": subject,
                "htmlBody": "<html><body>" + body + "</body></html>",
            }
        )

    return drafts, unmatched


def load_contacts(path):
    """Load a Ship-to name -> email map from a contacts workbook.
    Tolerant of column naming: uses the first column whose header contains
    'name' and the first whose header contains 'email'."""
    if not path or not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [str(c.value or "").strip().lower() for c in ws[1]]
    name_idx = next((i for i, h in enumerate(headers) if "name" in h), None)
    email_idx = next((i for i, h in enumerate(headers) if "email" in h or "mail" in h), None)
    contacts = {}
    if name_idx is None or email_idx is None:
        return contacts
    for row in ws.iter_rows(min_row=2, values_only=True):
        if name_idx < len(row) and email_idx < len(row):
            name = row[name_idx]
            email = row[email_idx]
            if name and email:
                contacts[norm_name(name)] = str(email).strip()
    return contacts


def main():
    report_path = sys.argv[1] if len(sys.argv) > 1 else None
    contacts_path = sys.argv[2] if len(sys.argv) > 2 else None
    if not report_path:
        print("Usage: parse_and_preview.py <report.xlsx> [contacts.xlsx]")
        sys.exit(1)

    wb = openpyxl.load_workbook(report_path, data_only=True)
    ws = wb["Consolidations"] if "Consolidations" in wb.sheetnames else wb.active

    blocks = parse_consolidations(ws)
    contacts = load_contacts(contacts_path)

    drafts, unmatched = build_drafts(blocks, contacts)

    out_dir = os.path.join(os.path.dirname(__file__), "..", "samples")
    os.makedirs(out_dir, exist_ok=True)

    with open(os.path.join(out_dir, "draft_emails.json"), "w") as f:
        json.dump(drafts, f, indent=2, default=str)

    for i, d in enumerate(drafts, 1):
        preview = (
            f"<!-- To: {d['to'] or '(unmatched - fill in)'} | CC: {d['cc']} -->\n"
            f"<!-- Subject: {d['subject']} -->\n"
            f"{d['htmlBody']}"
        )
        with open(os.path.join(out_dir, f"preview_{i}.html"), "w") as f:
            f.write(preview)

    print(f"Parsed {len(blocks)} consolidations.")
    print(f"Built {len(drafts)} draft email(s).")
    if contacts:
        print(f"Loaded {len(contacts)} contacts.")
    else:
        print("No contacts workbook supplied -> recipients left blank in previews.")
    if unmatched:
        print(f"Unmatched Ship-to names ({len(unmatched)}): {sorted(set(unmatched))}")
    for i, d in enumerate(drafts, 1):
        print(f"  Draft {i}: to={d['to'] or '(blank)'} | {d['subject']}")


if __name__ == "__main__":
    main()
