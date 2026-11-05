# Nodo KNIME: Python Script (2 entradas, 2 salidas)  — "Match NOVA vs SAP (trigger)"
# Entrada 0: tabla NOVA (columna 'po', salida del nodo 01)
# Entrada 1: tabla SAP  (columna indicada en config sap.po_column, p. ej. BSTKD)
# Salida 0: resumen (1 fila) con conteos y si se disparó
# Salida 1: detalle PO por PO (MATCH / SOLO_NOVA / SOLO_SAP)
# Flow vars de salida:
#   trigger_fired  -> 1/0
#   PortChoice     -> "top" si se dispara, "bottom" si no  (conéctalo al nodo IF Switch)
#   trigger_reason -> texto
import json
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

import knime.scripting.io as knio
import pandas as pd

LIB_DIR = knio.flow_variables["lib_dir"]
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
with open(knio.flow_variables["config_path"], encoding="utf-8") as fh:
    CONFIG = json.load(fh)

from nova_sap.po_matching import build_detail_rows, evaluate_trigger  # noqa: E402
from nova_sap.state import StateStore  # noqa: E402

m = CONFIG.get("matching", {})
nova = knio.input_tables[0].to_pandas()
sap = knio.input_tables[1].to_pandas()
sap_col = CONFIG["sap"]["po_column"]
if sap_col not in sap.columns:
    raise KeyError(f"La tabla SAP no tiene la columna '{sap_col}'. Columnas: {list(sap.columns)}")

zeros = m.get("strip_leading_zeros", True)
result = evaluate_trigger(nova["po"].tolist(), sap[sap_col].tolist(),
                          strip_leading_zeros=zeros, allow_empty=m.get("allow_empty_nova", False))

business_date = datetime.now(ZoneInfo(m.get("timezone", "America/Mexico_City"))).date().isoformat()
store = StateStore(CONFIG["state_db"])
fired = result.fired
reason = result.reason
if fired and m.get("fire_once_per_day", True) and store.already_fired(business_date):
    fired, reason = False, "Ya se disparó hoy; no se repite"
summary = result.summary() | {"trigger_fired": fired, "reason": reason, "business_date": business_date,
                              "nova_error": knio.flow_variables.get("nova_error", "")}
store.log_trigger(business_date, fired, summary | {"missing_in_sap": result.missing_in_sap[:500]})

knio.flow_variables["trigger_fired"] = int(fired)
knio.flow_variables["PortChoice"] = "top" if fired else "bottom"
knio.flow_variables["trigger_reason"] = reason
knio.flow_variables["business_date"] = business_date

knio.output_tables[0] = knio.Table.from_pandas(pd.DataFrame([summary]))
detail = pd.DataFrame(build_detail_rows(nova["po"].tolist(), sap[sap_col].tolist(), zeros),
                      columns=["po", "veces_nova", "veces_sap", "estatus"])
knio.output_tables[1] = knio.Table.from_pandas(detail)
