# Nodo KNIME: Python Script (0 entradas, 2 salidas)  — "Procesar respuestas de OMAs"
# Workflow separado, programado cada 5-10 min durante el horario de operación.
# Salida 0: log de mensajes procesados (aprobado / editado / rechazado / ignorado / error)
# Salida 1: estado actual de todas las solicitudes del día
import json
import os
import sys

import knime.scripting.io as knio
import pandas as pd

LIB_DIR = knio.flow_variables["lib_dir"]
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
with open(knio.flow_variables["config_path"], encoding="utf-8") as fh:
    CONFIG = json.load(fh)

from nova_sap.approvals import expire_old, process_replies  # noqa: E402
from nova_sap.graph_client import GraphMailClient  # noqa: E402
from nova_sap.state import StateStore  # noqa: E402

o = CONFIG["outlook"]
client = GraphMailClient(os.environ[o["tenant_id_env"]], os.environ[o["client_id_env"]],
                         os.environ[o["client_secret_env"]],
                         o["sender_mailbox"], o["approvals_mailbox"])
store = StateStore(CONFIG["state_db"])

log = process_replies(store, client, o["approvals_mailbox"], o.get("delegates"))
log += expire_old(store, client, o.get("expires_hours", 8))

cols = ["message_id", "from", "subject", "request_id", "decision", "action", "detail"]
knio.output_tables[0] = knio.Table.from_pandas(pd.DataFrame(log).reindex(columns=cols).astype(str))

status = pd.read_sql_query(
    "SELECT request_id, business_date, report_name, oma_email, to_addrs, subject, status, "
    "created_ts, decided_ts, decision_note FROM approval_requests ORDER BY created_ts DESC LIMIT 500",
    store.conn)
knio.output_tables[1] = knio.Table.from_pandas(status.astype(str))
