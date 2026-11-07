"""Motor de reportes WMT para el servicio REST de KNIME Hub.

Power Automate llama al servicio con:
    {"reporte": "Pricing Report", "turno": "AM", "fecha": "2026-11-09",
     "asunto_formato": "WM Turn Pricing {FECHA:MM.dd.yyyy}",
     "machote_eml_base64": "<.eml del machote, opcional>",
     "modo_prueba": true}

y recibe:
    {"ok": true, "error": "",
     "correos": [{"asunto": ..., "html": ..., "to": "", "resumen": ...,
                  "adjunto_nombre": "x.xlsx", "adjunto_base64": "..."}]}

Los datos de cada reporte llegan como DataFrames (consultas DB a SAP en KNIME).
Mientras no haya consultas (modo_prueba) se usan datos de ejemplo.
"""
from __future__ import annotations

import base64
import html as _html
import io
import json
import re
from datetime import date, datetime
from email import policy
from email.parser import BytesParser

import pandas as pd

REPORTES = ["BGJs", "Pricing Report", "OB Report", "Rounding", "Line Block", "Collective Check"]

# ------------------------------------------------------------------ marcadores
_FMT = [("dddd", "%A"), ("yyyy", "%Y"), ("yy", "%y"), ("MM", "%m"), ("dd", "%d"), ("HH", "%H"), ("mm", "%M")]


def formato_fecha(fecha: date, fmt_net: str) -> str:
    """Convierte un formato estilo Outlook/.NET (MM.dd.yyyy) a strftime."""
    out, i = "", 0
    while i < len(fmt_net):
        for net, py in _FMT:
            if fmt_net.startswith(net, i):
                out += py
                i += len(net)
                break
        else:
            out += fmt_net[i].replace("%", "%%")
            i += 1
    return fecha.strftime(out)


def aplicar_marcadores(texto: str, fecha: date, turno: str, valores: dict) -> str:
    if not texto:
        return texto
    texto = re.sub(r"\{FECHA:([^}]+)\}", lambda m: formato_fecha(fecha, m.group(1)), texto)
    fijos = {
        "FECHA": fecha.strftime("%m.%d.%Y"),
        "DIA_EN": fecha.strftime("%A"),
        "TURNO_EN": "Afternoon" if turno == "PM" else "Morning",
        "TURNO": turno,
    }
    for k, v in {**fijos, **valores}.items():
        texto = texto.replace("{" + k + "}", str(v))
    return texto


def limpiar_marcadores(texto: str) -> str:
    """Marcadores sin datos: tablas -> NONE (como la macro), el resto -> vacio."""
    if not texto:
        return texto
    texto = re.sub(r"\{TABLA_[A-Z0-9_]+\}", "<p><i>NONE</i></p>", texto)
    return re.sub(r"\{[A-Z][A-Z0-9_]*\}", "", texto)


# ------------------------------------------------------------------ html / excel
def tabla_html(df: pd.DataFrame, max_filas: int = 250) -> str:
    if df is None or df.empty:
        return "<p><i>No data.</i></p>"
    th = "".join(f"<th style='background:#4D93D9;color:#fff;padding:4px 8px;border:1px solid #9bb'>{_html.escape(str(c))}</th>"
                 for c in df.columns)
    filas = []
    for _, r in df.head(max_filas).iterrows():
        tds = "".join(f"<td style='padding:3px 8px;border:1px solid #ccc'>{_html.escape('' if pd.isna(v) else _fmt(v))}</td>"
                      for v in r.values)
        filas.append(f"<tr>{tds}</tr>")
    extra = f"<p style='color:#666'>Showing {max_filas} of {len(df)} rows (full list attached).</p>" if len(df) > max_filas else ""
    return ("<table style='border-collapse:collapse;font-family:Calibri,Arial;font-size:10pt'>"
            f"<tr>{th}</tr>{''.join(filas)}</table>{extra}")


def _fmt(v) -> str:
    if isinstance(v, float) and v.is_integer():
        return f"{int(v):,}"
    if isinstance(v, (int, float)):
        return f"{v:,}"
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return v.strftime("%m/%d/%Y")
    return str(v)


def excel_base64(hojas: dict[str, pd.DataFrame]) -> str:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nombre, df in hojas.items():
            df.to_excel(xw, sheet_name=nombre[:31], index=False)
    return base64.b64encode(buf.getvalue()).decode()


# ------------------------------------------------------------------ machote (.eml)
def cuerpo_machote(eml_b64: str | None) -> str | None:
    """HTML del machote: un correo diseñado en Outlook y descargado como .eml."""
    if not eml_b64:
        return None
    msg = BytesParser(policy=policy.default).parsebytes(base64.b64decode(eml_b64))
    parte = msg.get_body(preferencelist=("html", "plain"))
    if parte is None:
        return None
    contenido = parte.get_content()
    if parte.get_content_type() == "text/plain":
        contenido = "<div style='font-family:Calibri,Arial;font-size:11pt'>" + \
            _html.escape(contenido).replace("\n", "<br>") + "</div>"
    return contenido


def cuerpo_desde_texto(texto: str | None) -> str | None:
    """Cuerpo escrito en Excel -> HTML con el estilo de Outlook.
    Reglas: cada linea = parrafo; '- ' = vineta; **negrita**; [rojo]texto[/rojo];
    un marcador solo en su linea (p.ej. {TABLA_PIVOT}) se inserta tal cual."""
    if not texto or not str(texto).strip():
        return None

    def inline(t: str) -> str:
        t = _html.escape(t, quote=False)
        t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
        t = re.sub(r"\[rojo\](.+?)\[/rojo\]", r"<span style='color:red;font-weight:bold'>\1</span>", t, flags=re.I)
        return t

    partes, lista = [], []

    def cerrar_lista():
        if lista:
            partes.append("<ul style='margin:0 0 10px 0'>" + "".join(f"<li>{x}</li>" for x in lista) + "</ul>")
            lista.clear()

    for linea in str(texto).replace("\r\n", "\n").split("\n"):
        l = linea.strip()
        if l.startswith("- "):
            lista.append(inline(l[2:]))
            continue
        cerrar_lista()
        if not l:
            continue
        if re.fullmatch(r"\{[A-Z][A-Z0-9_]*\}", l):
            partes.append(l)
        else:
            partes.append(f"<p style='margin:0 0 10px 0'>{inline(l)}</p>")
    cerrar_lista()
    return "<div style='font-family:Calibri,Arial;font-size:11pt;color:#000'>" + "".join(partes) + "</div>"


def cuerpo_por_defecto(bloques: dict[str, str]) -> str:
    partes = ["<div style='font-family:Calibri,Arial;font-size:11pt'><p>Hi Team,</p>"]
    partes += [v for v in bloques.values() if v]
    partes.append("<p>Thanks,</p></div>")
    return "\n".join(partes)


# ------------------------------------------------------------------ datos de prueba
def datos_prueba(reporte: str) -> dict[str, pd.DataFrame]:
    if reporte == "BGJs":
        return {"Merch": pd.DataFrame({"Plant": ["B360", "B360", "B275"], "Ship-to name1": ["WM DC 6006", "WM DC 6006", "WM DC 7038"],
                                       "Document number": ["0101000001", "0101000002", "0101000003"],
                                       "OMA": ["oma1@pg.com", "oma1@pg.com", "oma2@pg.com"]})}
    if reporte == "OB Report":
        return {"ZVOMH": pd.DataFrame({"Document number": ["0101000010", "0101000011"], "Delivery block": ["09", "LU"],
                                       "Ship-to Your Reference": ["POS REPLEN", "WMT 6006"], "Created On": ["11/06/2026", "11/06/2026"]})}
    return {"Pivot": pd.DataFrame({"Plant": ["B360", "B275"], "Category": ["Laundry", "Paper"], "Amount": [1200, 800]})}


# ------------------------------------------------------------------ construccion de correos
def construir_correos(sol: dict, datos: dict[str, pd.DataFrame] | None = None) -> list[dict]:
    reporte = sol.get("reporte", "")
    if reporte not in REPORTES:
        raise ValueError(f"Reporte desconocido: {reporte!r}. Validos: {REPORTES}")
    turno = (sol.get("turno") or "AM").upper()
    fecha = datetime.strptime(sol["fecha"], "%Y-%m-%d").date() if sol.get("fecha") else date.today()
    if datos is None:
        datos = datos_prueba(reporte)
    prueba = bool(sol.get("modo_prueba"))
    machote = cuerpo_desde_texto(sol.get("cuerpo_texto")) or cuerpo_machote(sol.get("machote_eml_base64"))
    asunto_fmt = sol.get("asunto_formato") or (reporte + " {FECHA:MM.dd.yyyy}")
    etiqueta = "[PRUEBA] " if prueba else ""
    fecha_txt = fecha.strftime("%m.%d.%Y")

    def un_correo(hojas: dict[str, pd.DataFrame], extra: dict, to: str = "") -> dict:
        tablas = {f"TABLA_{n.upper().replace(' ', '_')}": tabla_html(df) for n, df in hojas.items()}
        primera = next(iter(tablas.values()), "")
        valores = {"TABLA_PIVOT": primera, **tablas, **extra}
        cuerpo = machote if machote else cuerpo_por_defecto({"tabla": primera})
        html = limpiar_marcadores(aplicar_marcadores(cuerpo, fecha, turno, valores))
        asunto = etiqueta + limpiar_marcadores(aplicar_marcadores(asunto_fmt, fecha, turno, valores))
        filas = sum(len(df) for df in hojas.values())
        return {"asunto": asunto, "html": html, "to": to,
                "resumen": f"{reporte} ({turno}) - {filas} filas" + (" - DATOS DE PRUEBA" if prueba else ""),
                "adjunto_nombre": f"{reporte} {fecha_txt}.xlsx", "adjunto_base64": excel_base64(hojas)}

    if reporte == "BGJs":                       # un correo por OMA con sus plantas
        df = next(iter(datos.values()))
        correos = []
        for oma, grupo in df.groupby("OMA", sort=True):
            plantas = ", ".join(sorted(grupo["Plant"].astype(str).unique()))
            nombre_oma = str(grupo["OMA_Nombre"].iloc[0]) if "OMA_Nombre" in grupo else str(oma)
            c = un_correo({"Merch": grupo.drop(columns=[x for x in ("OMA", "OMA_Nombre") if x in grupo])},
                          {"PLANTAS": plantas, "OMA": nombre_oma}, to=str(oma))
            c["plantas"] = plantas
            correos.append(c)
        return correos

    extra = {}
    if reporte == "OB Report":
        z = datos.get("ZVOMH", pd.DataFrame())
        n09 = int(((z.get("Delivery block", pd.Series(dtype=str)).astype(str) == "09") &
                   z.get("Ship-to Your Reference", pd.Series(dtype=str)).astype(str).str.contains("POS", case=False)).sum()) if not z.empty else 0
        extra["ALERTA_09"] = (f"<p style='color:#c00;font-weight:bold'>THERE ARE {n09} 09 POS REPLEN BLOCKS WITHIN OPP WINDOW "
                              "THAT NEED YOUR ATTENTION</p>") if n09 else ""
    return [un_correo(datos, extra)]


# ------------------------------------------------------------------ configuracion (WMT_Config.xlsx)
def leer_config(xlsx_b64: str | None) -> dict:
    """Lee WMT_Config.xlsx (hojas Configuracion, Disparadores, Contactos, Plantillas, OMA_Plantas)."""
    if not xlsx_b64:
        return {}
    hojas = pd.read_excel(io.BytesIO(base64.b64decode(xlsx_b64)), sheet_name=None, dtype=str)
    cfg = {n.strip(): df.fillna("").apply(lambda c: c.str.strip()) for n, df in hojas.items()}
    general = {}
    if "Configuracion" in cfg:
        for _, r in cfg["Configuracion"].iterrows():
            general[r.iloc[0]] = r.iloc[1]
    cfg["_general"] = general
    return cfg


def _si(v) -> bool:
    return str(v).strip().lower() in ("si", "sí", "yes", "true", "1")


def _fila(cfg: dict, hoja: str, reporte: str) -> dict:
    df = cfg.get(hoja)
    if df is None or df.empty or "Reporte" not in df:
        return {}
    m = df[df["Reporte"].str.lower() == reporte.lower()]
    return m.iloc[0].to_dict() if len(m) else {}


def _correos(s: str) -> list[str]:
    return [x.strip() for x in re.split(r"[;,]", s or "") if x.strip()]


def _graph(asunto, html, to, cc, adj_nombre, adj_b64) -> dict:
    msg = {"subject": asunto, "body": {"contentType": "HTML", "content": html},
           "toRecipients": [{"emailAddress": {"address": a}} for a in _correos(to)],
           "ccRecipients": [{"emailAddress": {"address": a}} for a in _correos(cc)]}
    if adj_b64:
        msg["attachments"] = [{"@odata.type": "#microsoft.graph.fileAttachment", "name": adj_nombre, "contentBytes": adj_b64}]
    return msg


def plan(sol: dict, cfg: dict) -> dict:
    """Que reportes corren para el correo disparador (ALL IN / ALL DONE)."""
    disp = (sol.get("disparador") or "").strip().upper()
    g = cfg.get("_general", {})
    df = cfg.get("Disparadores", pd.DataFrame())
    fila = df[(df["Asunto exacto"].str.upper() == disp) & (df["Activo"].map(_si) | (df["Activo"] == ""))] if not df.empty else df
    if fila is None or fila.empty:
        return {"ok": True, "error": "", "reportes": [], "aprobador_default": g.get("Aprobador por defecto", "")}
    r = fila.iloc[0]
    reportes = []
    for rep in _correos(r["Reportes que corre"]):
        pl = _fila(cfg, "Plantillas", rep)
        reportes.append({"reporte": rep, "turno": (r.get("Turno OB") or "AM").upper(),
                         "machote": pl.get("Machote (.eml)", "")})
    return {"ok": True, "error": "", "reportes": reportes, "aprobador_default": g.get("Aprobador por defecto", "")}


def resolver(sol: dict, cfg: dict, correos: list[dict]) -> list[dict]:
    """Agrega To/CC/aprobadores/remitente desde la configuracion y arma el cuerpo de Graph."""
    reporte = sol.get("reporte", "")
    g = cfg.get("_general", {})
    prueba = _si(g.get("Modo prueba", "Si")) if cfg else bool(sol.get("modo_prueba"))
    aprob_def = g.get("Aprobador por defecto", "")
    ct = _fila(cfg, "Contactos", reporte)
    oma = cfg.get("OMA_Plantas", pd.DataFrame())
    out = []
    for c in correos:
        to = c.get("to") or ct.get("To", "")
        if "{OMA_POR_PLANTA}" in to and not oma.empty:
            plantas = [p.strip() for p in c.get("plantas", "").split(",") if p.strip()]
            to = "; ".join(sorted(set(oma[oma["Planta"].isin(plantas)]["Correo OMA"]) - {""}))
        cc = ct.get("CC", "")
        aprob = ct.get("Aprobadores") or aprob_def
        if prueba:
            to, cc = aprob or aprob_def, ""
            c["asunto"] = c["asunto"] if c["asunto"].startswith("[PRUEBA]") else "[PRUEBA] " + c["asunto"]
        enviar = (not ct) or ct.get("Activo", "Si") == "" or _si(ct.get("Activo", "Si"))
        gm = _graph(c["asunto"], c["html"], to, cc, c.get("adjunto_nombre", ""), c.get("adjunto_base64", ""))
        out.append({**c, "reporte": reporte, "to": to, "cc": cc, "aprobadores": aprob, "enviar": bool(enviar and to),
                    "de": ct.get("Enviar como (From)", g.get("Buzon de envio", "")),
                    "graph_mensaje": json.dumps(gm, ensure_ascii=False),
                    "graph_sendmail": json.dumps({"message": gm, "saveToSentItems": True}, ensure_ascii=False)})
    return out


def responder(sol: dict, datos: dict[str, pd.DataFrame] | None = None) -> dict:
    try:
        cfg = leer_config(sol.get("config_xlsx_base64"))
        if sol.get("accion") == "plan":
            return plan(sol, cfg)
        if cfg:
            pl = _fila(cfg, "Plantillas", sol.get("reporte", ""))
            sol = {**sol, "asunto_formato": sol.get("asunto_formato") or pl.get("Asunto (formato exacto)", ""),
                   "cuerpo_texto": sol.get("cuerpo_texto") or pl.get("Cuerpo del correo", ""),
                   "modo_prueba": _si(cfg["_general"].get("Modo prueba", "Si"))}
        correos = construir_correos(sol, datos)
        return {"ok": True, "error": "", "correos": resolver(sol, cfg, correos)}
    except Exception as exc:  # noqa: BLE001 - el error viaja a Power Automate
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}", "correos": [], "reportes": []}
