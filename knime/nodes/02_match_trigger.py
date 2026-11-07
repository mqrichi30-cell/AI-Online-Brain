# Nodo KNIME: Python Script (2 entradas, 2 salidas)  — "Match NOVA vs SAP (trigger)"
# Entrada 0: tabla NOVA   |   Entrada 1: tabla SAP
# Salida 0: resumen (1 fila) con conteos y si se disparó
# Salida 1: detalle PO por PO (MATCH / SOLO_NOVA / SOLO_SAP)
#
# Flow variables de entrada (todas opcionales):
#   lib_dir, config_path -> si no vienen, se usan data/lib y data/config.json del workflow
#   nova_po_column, sap_po_column -> nombre de la columna de PO; vacío = detectar sola
# Flow variables de salida:
#   trigger_fired (1/0), PortChoice ("top" si se dispara / "bottom"), trigger_reason, business_date
import json
import os
import re
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

import knime.scripting.io as knio
import pandas as pd


def _fv(name, default=""):
    try:
        value = knio.flow_variables[name]
    except KeyError:
        return default
    return value if value not in (None, "") else default


DATA_DIR = None
try:
    DATA_DIR = knio.get_workflow_data_area_dir()
except Exception:  # noqa: BLE001 - versiones de KNIME sin data area
    pass
LIB_DIR = _fv("lib_dir") or os.path.join(DATA_DIR, "lib")
CONFIG_PATH = _fv("config_path") or os.path.join(DATA_DIR, "config.json")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
with open(CONFIG_PATH, encoding="utf-8") as fh:
    CONFIG = json.load(fh)

from nova_sap.po_matching import build_detail_rows, evaluate_trigger  # noqa: E402
from nova_sap.state import StateStore  # noqa: E402

PO_NAME_HINTS = re.compile(r"^(po|po ?(number|nbr|num|no|#)|purchase ?order.*|bstkd|bstnk|"
                           r"customer ?(po|reference).*|referencia.*cliente|pedido ?cliente|orden ?de ?compra)$",
                           re.IGNORECASE)


def pick_column(df: pd.DataFrame, wanted: str, side: str) -> str:
    if wanted:
        if wanted not in df.columns:
            raise KeyError(f"{side}: no existe la columna '{wanted}'. Columnas: {list(df.columns)}")
        return wanted
    candidates = [c for c in df.columns if PO_NAME_HINTS.match(str(c).strip())]
    if len(candidates) == 1:
        return candidates[0]
    raise KeyError(
        f"{side}: no pude detectar la columna de PO (candidatas: {candidates}). Escribe su nombre en la "
        f"variable '{side.lower()}_po_column' del nodo Variable Creator. Columnas: {list(df.columns)}")


m = CONFIG.get("matching", {})
nova = knio.input_tables[0].to_pandas()
sap = knio.input_tables[1].to_pandas()
nova_col = pick_column(nova, _fv("nova_po_column", "po" if "po" in nova.columns else ""), "NOVA")
sap_col = pick_column(sap, _fv("sap_po_column", CONFIG.get("sap", {}).get("po_column", "")
                             if CONFIG.get("sap", {}).get("po_column") in sap.columns else ""), "SAP")
nova_pos, sap_pos = nova[nova_col].tolist(), sap[sap_col].tolist()

zeros = m.get("strip_leading_zeros", True)
result = evaluate_trigger(nova_pos, sap_pos, strip_leading_zeros=zeros,
                          allow_empty=m.get("allow_empty_nova", False))

try:
    business_date = datetime.now(ZoneInfo(m.get("timezone", "America/Mexico_City"))).date().isoformat()
except Exception:  # noqa: BLE001 - sin base de zonas horarias: usar la hora local de la PC
    business_date = datetime.now().date().isoformat()
state_db = CONFIG["state_db"]
if not os.path.isabs(state_db):
    state_db = os.path.join(os.path.dirname(os.path.abspath(CONFIG_PATH)), state_db)
store = StateStore(state_db)
fired, reason = result.fired, result.reason
if fired and m.get("fire_once_per_day", True) and store.already_fired(business_date):
    fired, reason = False, "Ya se disparó hoy; no se repite"
summary = result.summary() | {"trigger_fired": fired, "reason": reason, "business_date": business_date,
                              "nova_po_column": nova_col, "sap_po_column": sap_col,
                              "nova_error": _fv("nova_error")}
store.log_trigger(business_date, fired, summary | {"missing_in_sap": result.missing_in_sap[:500]})

knio.flow_variables["trigger_fired"] = int(fired)
knio.flow_variables["PortChoice"] = "top" if fired else "bottom"
knio.flow_variables["trigger_reason"] = reason
knio.flow_variables["business_date"] = business_date

knio.output_tables[0] = knio.Table.from_pandas(pd.DataFrame([summary]))
detail = pd.DataFrame(build_detail_rows(nova_pos, sap_pos, zeros),
                      columns=["po", "veces_nova", "veces_sap", "estatus"])
knio.output_tables[1] = knio.Table.from_pandas(detail)
