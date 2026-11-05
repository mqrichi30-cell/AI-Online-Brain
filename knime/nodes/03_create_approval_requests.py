# Nodo KNIME: Python Script (1 entrada, 1 salida)  — "Crear solicitudes de aprobación"
# Entrada 0: un renglón por correo que generan los reportes. Columnas:
#   report_name | oma_email | to | cc | subject | body_html | attachment_paths
#   (to / cc / attachment_paths separados por ';')
# Salida 0: request_id, oma, subject, status
# Credenciales: variables de entorno del usuario que corre KNIME, con los nombres
# indicados en config.outlook.*_env (tenant, client id, client secret).
import json
import mimetypes
import os
import sys
from pathlib import Path

import knime.scripting.io as knio
import pandas as pd

LIB_DIR = knio.flow_variables["lib_dir"]
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
with open(knio.flow_variables["config_path"], encoding="utf-8") as fh:
    CONFIG = json.load(fh)

from nova_sap.approvals import OutgoingEmail, create_approval_requests  # noqa: E402
from nova_sap.graph_client import GraphMailClient  # noqa: E402
from nova_sap.state import StateStore  # noqa: E402

o = CONFIG["outlook"]


def graph_client() -> GraphMailClient:
    return GraphMailClient(os.environ[o["tenant_id_env"]], os.environ[o["client_id_env"]],
                           os.environ[o["client_secret_env"]],
                           o["sender_mailbox"], o["approvals_mailbox"])


def split(v) -> list[str]:
    return [x.strip() for x in str(v or "").split(";") if x.strip() and x.strip().lower() != "nan"]


def load_attachments(paths: str):
    out = []
    for p in split(paths):
        ctype = mimetypes.guess_type(p)[0] or "application/octet-stream"
        out.append((Path(p).name, Path(p).read_bytes(), ctype))
    return out


df = knio.input_tables[0].to_pandas()
mails = [
    OutgoingEmail(report_name=r["report_name"], oma_email=r["oma_email"].strip(), to=split(r["to"]),
                  cc=split(r.get("cc")), subject=r["subject"], body_html=r["body_html"],
                  attachments=load_attachments(r.get("attachment_paths")))
    for _, r in df.iterrows()
]
result = create_approval_requests(
    StateStore(CONFIG["state_db"]), graph_client(), mails,
    business_date=knio.flow_variables.get("business_date", ""),
    approvals_mailbox=o["approvals_mailbox"], expires_hours=o.get("expires_hours", 8))

knio.output_tables[0] = knio.Table.from_pandas(
    pd.DataFrame(result, columns=["request_id", "oma", "subject", "status"]))
