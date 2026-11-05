# Nodo KNIME: Python Script (0 entradas, 1 salida)  — "Extraer órdenes NOVA"
# Salida 0: tabla NOVA con columna 'po' + columnas originales (todo texto).
# Flow vars de salida: nova_ok (1/0), nova_error (texto)
import json
import sys

import knime.scripting.io as knio
import pandas as pd

LIB_DIR = knio.flow_variables["lib_dir"]
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
with open(knio.flow_variables["config_path"], encoding="utf-8") as fh:
    CONFIG = json.load(fh)

from nova_sap.nova_extract import NovaLoginRequired, extract_nova_orders  # noqa: E402

try:
    df = extract_nova_orders(CONFIG, headless=True).astype(str)
    knio.flow_variables["nova_ok"] = 1
    knio.flow_variables["nova_error"] = ""
except NovaLoginRequired as exc:
    # No tronar el workflow: el match verá NOVA vacío y NO disparará; se avisa aparte.
    df = pd.DataFrame({"po": pd.Series(dtype=str)})
    knio.flow_variables["nova_ok"] = 0
    knio.flow_variables["nova_error"] = f"LOGIN_REQUERIDO: {exc}"

knio.output_tables[0] = knio.Table.from_pandas(df)
