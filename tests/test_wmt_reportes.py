import base64
from email.message import EmailMessage

import openpyxl
import io

from nova_sap.wmt_reportes import aplicar_marcadores, formato_fecha, responder
from datetime import date

D = date(2026, 11, 7)


def test_formatos_de_asunto_iguales_a_las_macros():
    assert aplicar_marcadores("{DIA_EN} {TURNO_EN} Order Block Report {FECHA:MM.dd.yy}", D, "AM", {}) == \
        "Saturday Morning Order Block Report 11.07.26"
    assert aplicar_marcadores("{DIA_EN} {TURNO_EN} Order Block Report {FECHA:MM.dd.yy}", D, "PM", {}) == \
        "Saturday Afternoon Order Block Report 11.07.26"
    assert aplicar_marcadores("WM Turn Pricing {FECHA:MM.dd.yyyy}", D, "AM", {}) == "WM Turn Pricing 11.07.2026"
    assert formato_fecha(D, "dddd MM/dd") == "Saturday 11/07"


def test_pricing_con_machote_eml():
    m = EmailMessage()
    m["Subject"] = "machote"
    m.set_content("x")
    m.add_alternative("<p>Hola equipo {FECHA:MM.dd.yyyy}</p>{TABLA_PIVOT}<p>Saludos</p>", subtype="html")
    r = responder({"reporte": "Pricing Report", "turno": "AM", "fecha": "2026-11-07",
                   "asunto_formato": "WM Turn Pricing {FECHA:MM.dd.yyyy}",
                   "machote_eml_base64": base64.b64encode(m.as_bytes()).decode(), "modo_prueba": True})
    assert r["ok"], r["error"]
    c = r["correos"][0]
    assert c["asunto"] == "[PRUEBA] WM Turn Pricing 11.07.2026"
    assert "Hola equipo 11.07.2026" in c["html"] and "<table" in c["html"] and "Laundry" in c["html"]
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(c["adjunto_base64"])))
    assert wb.sheetnames == ["Pivot"] and c["adjunto_nombre"] == "Pricing Report 11.07.2026.xlsx"


def test_bgj_un_correo_por_oma():
    r = responder({"reporte": "BGJs", "fecha": "2026-11-07", "asunto_formato": "MC Distribution after BGJs - {PLANTAS}"})
    assert r["ok"]
    asuntos = {c["to"]: c["asunto"] for c in r["correos"]}
    assert asuntos == {"oma1@pg.com": "MC Distribution after BGJs - B360", "oma2@pg.com": "MC Distribution after BGJs - B275"}


def test_ob_alerta_09_y_reporte_invalido():
    r = responder({"reporte": "OB Report", "turno": "PM", "fecha": "2026-11-07"})
    assert "THERE ARE 1 09 POS REPLEN" in r["correos"][0]["html"] or "{ALERTA_09}" not in r["correos"][0]["html"]
    bad = responder({"reporte": "Nada"})
    assert not bad["ok"] and "Reporte desconocido" in bad["error"]


def _cfg_b64(modo_prueba="Si"):
    import openpyxl as ox
    wb = ox.Workbook(); ws = wb.active; ws.title = "Configuracion"
    for r in [["Parametro", "Valor"], ["Modo prueba", modo_prueba], ["Aprobador por defecto", "yo@pg.com"],
              ["Buzon de envio", "buzon@pg.com"]]:
        ws.append(r)
    d = wb.create_sheet("Disparadores")
    for r in [["Asunto exacto", "Reportes que corre", "Turno OB", "Activo"], ["ALL IN", "Pricing Report; BGJs", "AM", "Si"],
              ["ALL DONE", "OB Report", "PM", "Si"]]:
        d.append(r)
    c = wb.create_sheet("Contactos")
    for r in [["Reporte", "Activo", "To", "CC", "Enviar como (From)", "Aprobadores"],
              ["Pricing Report", "Si", "dl@pg.com", "cc1@pg.com; cc2@pg.com", "buzon@pg.com", ""],
              ["BGJs", "Si", "{OMA_POR_PLANTA}", "", "buzon@pg.com", "jefe@pg.com"]]:
        c.append(r)
    p = wb.create_sheet("Plantillas")
    for r in [["Reporte", "Asunto (formato exacto)", "Machote (.eml)"], ["Pricing Report", "WM Turn Pricing {FECHA:MM.dd.yyyy}", "Pricing Report.eml"],
              ["BGJs", "MC Distribution after BGJs - {PLANTAS}", ""]]:
        p.append(r)
    o = wb.create_sheet("OMA_Plantas")
    for r in [["Planta", "OMA", "Correo OMA", "Activo"], ["B360", "Ana", "ana@pg.com", "Si"], ["B275", "Luis", "luis@pg.com", "Si"]]:
        o.append(r)
    b = io.BytesIO(); wb.save(b)
    return base64.b64encode(b.getvalue()).decode()


def test_plan_por_disparador():
    import json as js
    r = responder({"accion": "plan", "disparador": "all in", "config_xlsx_base64": _cfg_b64()})
    assert r["ok"] and [x["reporte"] for x in r["reportes"]] == ["Pricing Report", "BGJs"]
    assert r["reportes"][0] == {"reporte": "Pricing Report", "turno": "AM", "machote": "Pricing Report.eml"}
    assert responder({"accion": "plan", "disparador": "RE: ALL IN", "config_xlsx_base64": _cfg_b64()})["reportes"] == []
    pm = responder({"accion": "plan", "disparador": "ALL DONE", "config_xlsx_base64": _cfg_b64()})
    assert pm["reportes"][0]["turno"] == "PM"


def test_reporte_resuelto_produccion_y_prueba():
    import json as js
    prod = responder({"accion": "reporte", "reporte": "Pricing Report", "turno": "AM", "fecha": "2026-11-07",
                      "config_xlsx_base64": _cfg_b64("No")})
    c = prod["correos"][0]
    assert c["asunto"] == "WM Turn Pricing 11.07.2026" and c["to"] == "dl@pg.com" and c["cc"] == "cc1@pg.com; cc2@pg.com"
    assert c["aprobadores"] == "yo@pg.com" and c["enviar"] and c["de"] == "buzon@pg.com"
    gm = js.loads(c["graph_sendmail"])["message"]
    assert [x["emailAddress"]["address"] for x in gm["ccRecipients"]] == ["cc1@pg.com", "cc2@pg.com"]
    assert gm["attachments"][0]["name"] == "Pricing Report 11.07.2026.xlsx"

    prueba = responder({"accion": "reporte", "reporte": "Pricing Report", "fecha": "2026-11-07", "config_xlsx_base64": _cfg_b64("Si")})
    c = prueba["correos"][0]
    assert c["to"] == "yo@pg.com" and c["cc"] == "" and c["asunto"].startswith("[PRUEBA] ")


def test_bgj_resuelve_oma_por_planta():
    r = responder({"accion": "reporte", "reporte": "BGJs", "fecha": "2026-11-07", "config_xlsx_base64": _cfg_b64("No")})
    tos = sorted((c["to"], c["aprobadores"]) for c in r["correos"])
    assert tos == [("oma1@pg.com", "jefe@pg.com"), ("oma2@pg.com", "jefe@pg.com")]
