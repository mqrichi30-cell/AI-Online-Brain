"""Estado persistente (SQLite): log del trigger y solicitudes de aprobación.

Estados de una solicitud:
  PENDING        -> esperando respuesta del OMA
  APPROVED_SENT  -> OMA aprobó, KNIME envió el correo original
  EDITED_SENT    -> OMA aprobó con versión editada (adjunta) o la envió él mismo
  REJECTED       -> OMA rechazó
  EXPIRED        -> sin respuesta antes del límite
  ERROR          -> falló el envío
"""
from __future__ import annotations

import json
import sqlite3
import uuid
from datetime import datetime, timezone

SCHEMA = """
CREATE TABLE IF NOT EXISTS trigger_log (
    run_ts TEXT NOT NULL,
    business_date TEXT NOT NULL,
    fired INTEGER NOT NULL,
    summary_json TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS approval_requests (
    request_id TEXT PRIMARY KEY,
    business_date TEXT NOT NULL,
    report_name TEXT,
    oma_email TEXT NOT NULL,
    to_addrs TEXT NOT NULL,
    cc_addrs TEXT,
    subject TEXT NOT NULL,
    proposal_mime BLOB NOT NULL,
    status TEXT NOT NULL,
    created_ts TEXT NOT NULL,
    decided_ts TEXT,
    decision_note TEXT,
    approval_message_id TEXT
);
"""

FINAL_STATES = {"APPROVED_SENT", "EDITED_SENT", "REJECTED", "EXPIRED"}


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class StateStore:
    def __init__(self, path: str):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)

    # ---- trigger ----
    def already_fired(self, business_date: str) -> bool:
        cur = self.conn.execute(
            "SELECT 1 FROM trigger_log WHERE business_date=? AND fired=1 LIMIT 1", (business_date,)
        )
        return cur.fetchone() is not None

    def log_trigger(self, business_date: str, fired: bool, summary: dict) -> None:
        self.conn.execute(
            "INSERT INTO trigger_log VALUES (?,?,?,?)",
            (now_iso(), business_date, int(fired), json.dumps(summary, default=str)),
        )
        self.conn.commit()

    # ---- approvals ----
    @staticmethod
    def new_request_id() -> str:
        return "REQ-" + uuid.uuid4().hex[:8].upper()

    def create_request(self, *, request_id, business_date, report_name, oma_email, to_addrs,
                       cc_addrs, subject, proposal_mime: bytes, approval_message_id=None) -> None:
        self.conn.execute(
            "INSERT INTO approval_requests (request_id,business_date,report_name,oma_email,to_addrs,"
            "cc_addrs,subject,proposal_mime,status,created_ts,approval_message_id) "
            "VALUES (?,?,?,?,?,?,?,?, 'PENDING', ?, ?)",
            (request_id, business_date, report_name, oma_email.lower(), to_addrs, cc_addrs or "",
             subject, proposal_mime, now_iso(), approval_message_id),
        )
        self.conn.commit()

    def get(self, request_id: str):
        return self.conn.execute(
            "SELECT * FROM approval_requests WHERE request_id=?", (request_id,)
        ).fetchone()

    def pending(self):
        return self.conn.execute(
            "SELECT * FROM approval_requests WHERE status='PENDING' ORDER BY created_ts"
        ).fetchall()

    def set_status(self, request_id: str, status: str, note: str = "") -> bool:
        """Cambia estado sólo si sigue PENDING (evita doble envío). Devuelve True si cambió."""
        cur = self.conn.execute(
            "UPDATE approval_requests SET status=?, decided_ts=?, decision_note=? "
            "WHERE request_id=? AND status='PENDING'",
            (status, now_iso(), note[:2000], request_id),
        )
        self.conn.commit()
        return cur.rowcount == 1

    def force_status(self, request_id: str, status: str, note: str = "") -> None:
        self.conn.execute(
            "UPDATE approval_requests SET status=?, decided_ts=?, decision_note=? WHERE request_id=?",
            (status, now_iso(), note[:2000], request_id),
        )
        self.conn.commit()
