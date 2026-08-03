#!/usr/bin/env python3
"""
Ejecuta localmente la misma logica que el Office Script
ProcessWakefern11s_AUTOPUSH y que el flow
"Wakefern - 11s block - Autopush V14 WHOLESALE 2 ALERTS".

Sirve para dos cosas:
  1. Ejecutar la automatizacion cuando Power Automate no corrio,
     y obtener los dos correos listos para enviar.
  2. Verificar contra un .msg real que el script devuelve lo esperado,
     antes de volver a confiar en el flow.

Uso:
    python3 tools/run_local.py <archivo.msg|archivo.xlsx> [--today YYYY-MM-DD]
                               [--outdir output]

El .msg se lee sin dependencias externas de Outlook: solo olefile.
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import os
import re
import sys

# --- Espejo exacto de la configuracion del Office Script -------------------
SHEET_NAME = "Alerts"
CUSTOMER_MATCH = "WAKEFERN"
BLOCK_CODE = "11"
GI_WINDOW_DAYS = 1
PUSH_GI_BUSINESS_DAYS_AHEAD = 2
RDD_OFFSET_BUSINESS_DAYS = 1
TZ_OFFSET_HOURS = -6
HOLIDAYS: set[str] = set()

# --- Destinatarios, identicos a los del flow -------------------------------
OSS_TO = ["ossgenai.im@pg.com"]
CUSTOMER_TO = [
    "steve.salotti@wakefern.com",
    "Mark.Kielczynski@wakefern.com",
    "grocery_special_po_group@wakefern.com",
    "anne.mucchiello@wakefern.com",
]
CUSTOMER_CC = ["jackson.vs@pg.com"]
SENDER_SIGNATURE = "pgcustservw2.im@pg.com<br>NA Order Management | Regional"


# ===========================================================================
# Extraccion del adjunto .xlsx desde el .msg
# ===========================================================================
def xlsx_from_msg(msg_path: str) -> tuple[bytes, str]:
    """Devuelve (bytes del primer adjunto .xlsx no inline, nombre)."""
    try:
        import olefile
    except ImportError:
        sys.exit("Falta olefile. Instalar con:  pip install olefile")

    ole = olefile.OleFileIO(msg_path)

    def stream(path: str) -> str:
        try:
            return ole.openstream(path).read().decode("utf-16-le", "ignore")
        except Exception:
            return ""

    # Los adjuntos viven en __attach_version1.0_#XXXXXXXX
    prefixes = sorted(
        {e[0] for e in ole.listdir() if e[0].startswith("__attach_version1.0_")}
    )

    for prefix in prefixes:
        name = stream(f"{prefix}/__substg1.0_3707001F") or stream(
            f"{prefix}/__substg1.0_3001001F"
        )
        if not name.lower().endswith(".xlsx"):
            continue
        try:
            data = ole.openstream(f"{prefix}/__substg1.0_37010102").read()
        except Exception:
            continue
        if data[:4] == b"PK\x03\x04":
            return data, name

    sys.exit("No se encontro ningun adjunto .xlsx dentro del .msg.")


def msg_subject(msg_path: str) -> str:
    try:
        import olefile

        ole = olefile.OleFileIO(msg_path)
        return ole.openstream("__substg1.0_0037001F").read().decode("utf-16-le", "ignore")
    except Exception:
        return ""


# ===========================================================================
# Fechas -- espejo del Office Script
# ===========================================================================
def is_business_day(d: dt.date) -> bool:
    return d.weekday() < 5 and d.isoformat() not in HOLIDAYS


def add_business_days(d: dt.date, days: int) -> dt.date:
    result = d
    remaining = days
    while remaining > 0:
        result += dt.timedelta(days=1)
        if is_business_day(result):
            remaining -= 1
    return result


def format_long_date(d: dt.date) -> str:
    # %-d no es portable en todas las plataformas; se arma a mano.
    weekday = [
        "Monday", "Tuesday", "Wednesday", "Thursday",
        "Friday", "Saturday", "Sunday",
    ][d.weekday()]
    month = [
        "January", "February", "March", "April", "May", "June", "July",
        "August", "September", "October", "November", "December",
    ][d.month - 1]
    return f"{weekday}, {month} {d.day}, {d.year}"


def business_today() -> dt.date:
    return (dt.datetime.utcnow() + dt.timedelta(hours=TZ_OFFSET_HOURS)).date()


def cell_to_iso(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        epoch = dt.date(1899, 12, 30)
        return (epoch + dt.timedelta(days=int(value))).isoformat()

    text = str(value).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", text)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if m:
        return f"{m.group(3)}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    return ""


def normalize_block(value) -> str:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{int(round(value)):02d}"
    return str(value).strip() if value is not None else ""


def find_column(header: list, name: str) -> int:
    wanted = name.lower().strip()
    for i, h in enumerate(header):
        if str(h or "").lower().strip() == wanted:
            return i
    loose = re.sub(r"[.\s]", "", wanted)
    for i, h in enumerate(header):
        if re.sub(r"[.\s]", "", str(h or "").lower()) == loose:
            return i
    return -1


# ===========================================================================
# Logica principal -- espejo de main() del Office Script
# ===========================================================================
def process(xlsx_path: str, today: dt.date) -> dict:
    try:
        import openpyxl
    except ImportError:
        sys.exit("Falta openpyxl. Instalar con:  pip install openpyxl")

    import warnings

    warnings.filterwarnings("ignore")

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        sys.exit(
            f"La hoja '{SHEET_NAME}' no existe. Hojas encontradas: {wb.sheetnames}"
        )

    rows = list(wb[SHEET_NAME].values)
    if len(rows) < 2:
        sys.exit("La hoja 'Alerts' no tiene filas de datos.")

    header = list(rows[0])
    col_name = find_column(header, "Sold to Name")
    col_db = find_column(header, "DB")
    col_so = find_column(header, "Doc. Number")
    col_po = find_column(header, "PO Number")
    col_gi = find_column(header, "Good Issue")

    missing = [
        label
        for label, idx in [
            ("Sold to Name", col_name), ("DB", col_db),
            ("Doc. Number", col_so), ("PO Number", col_po),
            ("Good Issue", col_gi),
        ]
        if idx < 0
    ]
    if missing:
        sys.exit(f"Faltan columnas en la hoja 'Alerts': {missing}")

    push_gi = add_business_days(today, PUSH_GI_BUSINESS_DAYS_AHEAD)
    new_rdd = add_business_days(push_gi, RDD_OFFSET_BUSINESS_DAYS)
    targets = {(today + dt.timedelta(days=n)).isoformat() for n in range(GI_WINDOW_DAYS + 1)}

    matched = []
    for row in rows[1:]:
        sold_to = str(row[col_name] or "").upper()
        if CUSTOMER_MATCH not in sold_to:
            continue
        if normalize_block(row[col_db]) != BLOCK_CODE:
            continue
        gi_iso = cell_to_iso(row[col_gi])
        if gi_iso not in targets:
            continue
        matched.append(
            {
                "soNumber": str(row[col_so] or "").strip(),
                "poNumber": str(row[col_po] or "").strip(),
                "giIso": gi_iso,
            }
        )

    matched.sort(key=lambda r: (r["giIso"], r["soNumber"]))

    oss_rows = "".join(
        f"<tr><td>{html.escape(m['soNumber'])}</td><td>{push_gi.isoformat()}</td></tr>"
        for m in matched
    )
    client_rows = "".join(
        f"<tr><td>{html.escape(m['poNumber'])}</td><td>{new_rdd.isoformat()}</td></tr>"
        for m in matched
    )

    return {
        "rowCount": len(matched),
        "ossHtmlRows": oss_rows,
        "clientHtmlRows": client_rows,
        "clientSentenceDate": format_long_date(new_rdd),
        "todayIso": today.isoformat(),
        "tomorrowIso": (today + dt.timedelta(days=1)).isoformat(),
        "pushGiIso": push_gi.isoformat(),
        "_matched": matched,
        "_newRddIso": new_rdd.isoformat(),
    }


# ===========================================================================
# Render de los dos correos -- textos identicos a los del flow
# ===========================================================================
def render_oss_email(result: dict) -> dict:
    body = (
        "Hello,<br><br>"
        "Please apply this change:<br><br>"
        "<table border=1 cellspacing=0 cellpadding=4>"
        "<tr><th>SO #</th><th>Push GI to</th></tr>"
        f"{result['ossHtmlRows']}"
        "</table><br><br>"
        f"Regards,<br>{SENDER_SIGNATURE}"
    )
    return {
        "to": OSS_TO,
        "cc": [],
        "subject": f"Wakefern RDD Change Request APPOINTMENT LATE - {result['todayIso']}",
        "bodyHtml": body,
    }


def render_customer_email(result: dict) -> dict:
    body = (
        "Hello team<br><br>"
        "These POs had been assigned with an appointment too late, so we need "
        "to push to the soonest RDD possible that is "
        f"{result['clientSentenceDate']}, could you please provide an "
        "appointment?<br><br>"
        "<table border=1 cellspacing=0 cellpadding=4>"
        "<tr><th>PO #</th><th>New RDD</th></tr>"
        f"{result['clientHtmlRows']}"
        "</table><br><br>"
        f"Regards,<br>{SENDER_SIGNATURE}"
    )
    return {
        "to": CUSTOMER_TO,
        "cc": CUSTOMER_CC,
        "subject": "APPOINTMENT NEEDED WAKEFERN",
        "bodyHtml": body,
    }


def write_preview(path: str, oss: dict, customer: dict, result: dict) -> None:
    def block(title: str, mail: dict) -> str:
        cc = (
            f"<div><b>CC:</b> {html.escape(', '.join(mail['cc']))}</div>"
            if mail["cc"]
            else ""
        )
        return f"""
  <section>
    <h2>{html.escape(title)}</h2>
    <div class="meta">
      <div><b>Para:</b> {html.escape(', '.join(mail['to']))}</div>
      {cc}
      <div><b>Asunto:</b> {html.escape(mail['subject'])}</div>
    </div>
    <div class="body">{mail['bodyHtml']}</div>
  </section>"""

    doc = f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Wakefern 11s Autopush - correos generados</title>
<style>
  body {{ font-family: Segoe UI, system-ui, sans-serif; margin: 2rem; max-width: 60rem; }}
  section {{ border: 1px solid #ccc; border-radius: 8px; padding: 1rem 1.25rem; margin-bottom: 1.5rem; }}
  h1 {{ font-size: 1.4rem; }}
  h2 {{ font-size: 1.1rem; margin-top: 0; }}
  .meta {{ background: #f5f5f5; padding: .75rem; border-radius: 6px; font-size: .9rem; margin-bottom: 1rem; }}
  .body {{ line-height: 1.5; }}
  table {{ border-collapse: collapse; }}
  td, th {{ padding: 4px 8px; }}
</style>
</head>
<body>
<h1>Wakefern - 11s block - Autopush &mdash; ejecucion {html.escape(result['todayIso'])}</h1>
<p><b>Ordenes encontradas:</b> {result['rowCount']} &nbsp;|&nbsp;
   <b>Push GI a:</b> {html.escape(result['pushGiIso'])} &nbsp;|&nbsp;
   <b>Nuevo RDD:</b> {html.escape(result['_newRddIso'])}</p>
{block('Correo 1 - OSS', oss)}
{block('Correo 2 - Cliente (Wakefern)', customer)}
</body>
</html>"""
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(doc)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="archivo .msg o .xlsx del reporte")
    parser.add_argument(
        "--today",
        help="fecha de proceso YYYY-MM-DD (por defecto: hoy en UTC-6)",
    )
    parser.add_argument("--outdir", default="output", help="carpeta de salida")
    args = parser.parse_args()

    today = (
        dt.date.fromisoformat(args.today) if args.today else business_today()
    )

    os.makedirs(args.outdir, exist_ok=True)

    if args.input.lower().endswith(".msg"):
        data, name = xlsx_from_msg(args.input)
        xlsx_path = os.path.join(args.outdir, name)
        with open(xlsx_path, "wb") as fh:
            fh.write(data)
        print(f"Adjunto extraido: {name} ({len(data)} bytes)")
        print(f"Asunto del correo: {msg_subject(args.input)}")
    else:
        xlsx_path = args.input

    result = process(xlsx_path, today)

    oss = render_oss_email(result)
    customer = render_customer_email(result)

    payload = {k: v for k, v in result.items() if not k.startswith("_")}

    with open(os.path.join(args.outdir, "script_result.json"), "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2)
    with open(os.path.join(args.outdir, "emails.json"), "w", encoding="utf-8") as fh:
        json.dump({"oss": oss, "customer": customer}, fh, indent=2)

    write_preview(os.path.join(args.outdir, "preview.html"), oss, customer, result)

    print()
    print(f"Fecha de proceso : {result['todayIso']} (manana: {result['tomorrowIso']})")
    print(f"Push GI a        : {result['pushGiIso']}")
    print(f"Nuevo RDD        : {result['_newRddIso']}  ({result['clientSentenceDate']})")
    print(f"Ordenes {CUSTOMER_MATCH} DB {BLOCK_CODE} con GI hoy/manana: {result['rowCount']}")
    for m in result["_matched"]:
        print(f"   SO {m['soNumber']}   PO {m['poNumber']}   GI {m['giIso']}")
    print()
    print(f"Salidas escritas en: {args.outdir}/")
    print("   script_result.json  <- lo que debe devolver el Office Script")
    print("   emails.json         <- los dos correos, listos para enviar")
    print("   preview.html        <- vista previa para revisar antes de enviar")


if __name__ == "__main__":
    main()
