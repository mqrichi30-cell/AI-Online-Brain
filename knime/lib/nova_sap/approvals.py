"""Flujo de aprobación por correo (Outlook / Microsoft 365).

1. create_approval_requests(): por cada correo que generen los reportes se arma la
   "propuesta" (.eml editable) y se manda al OMA un correo de solicitud con:
     - la propuesta adjunta (abre en Outlook como borrador editable, X-Unsent: 1)
     - botones APROBAR / RECHAZAR (mailto: con el ID de solicitud en el asunto)
2. process_replies(): lee el buzón de aprobaciones y aplica la decisión:
     - APROBAR sin adjunto           -> KNIME envía la propuesta original   (APPROVED_SENT)
     - APROBAR con correo adjunto    -> KNIME envía la versión editada      (EDITED_SENT)
     - El OMA abre el borrador, edita y lo envía él mismo; la copia BCC al
       buzón de aprobaciones cierra la solicitud                           (EDITED_SENT)
     - RECHAZAR [motivo]             -> no se envía nada                    (REJECTED)
   Sólo se aceptan decisiones del OMA asignado (o delegados configurados).
3. expire_old(): marca EXPIRED lo que no se contestó a tiempo y avisa al OMA.
"""
from __future__ import annotations

import html
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from typing import Iterable, Protocol
from urllib.parse import quote

from .state import StateStore

REQ_ID_RE = re.compile(r"\bREQ-[0-9A-F]{8}\b", re.IGNORECASE)
APPROVE_WORDS = ("APROBAR", "APROBADO", "APRUEBO", "APPROVE", "APPROVED")
REJECT_WORDS = ("RECHAZAR", "RECHAZADO", "RECHAZO", "REJECT", "REJECTED")
_REPLY_PREFIX = re.compile(r"^\s*((RE|RV|FW|FWD|TR)\s*:\s*)+", re.IGNORECASE)


# --------------------------------------------------------------------------- tipos
@dataclass
class OutgoingEmail:
    """Un correo que un reporte quiere enviar (contrato de la tabla de entrada en KNIME)."""
    report_name: str
    oma_email: str
    to: list[str]
    cc: list[str]
    subject: str
    body_html: str
    attachments: list[tuple[str, bytes, str]]  # (nombre, bytes, content-type)


@dataclass
class InboxMessage:
    id: str
    sender: str
    subject: str
    body_text: str
    to: list[str]
    has_attachments: bool = False


class MailClient(Protocol):
    def send_html(self, to: list[str], subject: str, body_html: str,
                  attachments: list[tuple[str, bytes, str]] | None = None) -> None: ...
    def send_mime(self, mime_bytes: bytes) -> None: ...
    def list_unprocessed(self) -> list[InboxMessage]: ...
    def get_email_attachments(self, message_id: str) -> list[bytes]: ...
    def mark_processed(self, message_id: str) -> None: ...


# --------------------------------------------------------------------------- MIME
def build_proposal_mime(mail: OutgoingEmail, approvals_mailbox: str | None = None) -> bytes:
    """Arma la propuesta como .eml. X-Unsent: 1 hace que Outlook la abra como borrador
    editable (con destinatarios y adjuntos ya puestos). Se agrega el buzón de
    aprobaciones en BCC para detectar si el OMA la envía directamente."""
    msg = EmailMessage(policy=policy.SMTP)
    msg["To"] = ", ".join(mail.to)
    if mail.cc:
        msg["Cc"] = ", ".join(mail.cc)
    if approvals_mailbox:
        msg["Bcc"] = approvals_mailbox
    msg["Subject"] = mail.subject
    msg["X-Unsent"] = "1"
    msg.set_content("Este correo requiere un cliente que soporte HTML.")
    msg.add_alternative(mail.body_html, subtype="html")
    for name, data, ctype in mail.attachments:
        maintype, _, subtype = (ctype or "application/octet-stream").partition("/")
        msg.add_attachment(data, maintype=maintype, subtype=subtype or "octet-stream", filename=name)
    return msg.as_bytes()


def prepare_for_sending(mime_bytes: bytes, approvals_mailbox: str | None = None) -> bytes:
    """Limpia la propuesta (o la versión editada) antes de que KNIME la envíe:
    quita X-Unsent, From/Sender (Graph pone el buzón emisor), ids previos y el BCC
    al buzón de aprobaciones."""
    msg = BytesParser(policy=policy.SMTP).parsebytes(mime_bytes)
    for h in ("X-Unsent", "From", "Sender", "Message-ID", "Date", "In-Reply-To", "References"):
        del msg[h]
    if approvals_mailbox and msg.get("Bcc"):
        keep = [a.strip() for a in str(msg["Bcc"]).split(",")
                if a.strip() and approvals_mailbox.lower() not in a.lower()]
        del msg["Bcc"]
        if keep:
            msg["Bcc"] = ", ".join(keep)
    if not msg.get("To"):
        raise ValueError("El correo a enviar no tiene destinatarios (To)")
    return msg.as_bytes()


def build_request_html(request_id: str, mail: OutgoingEmail, approvals_mailbox: str,
                       expires_hours: int) -> str:
    def mailto(word: str) -> str:
        subj = quote(f"{word} [{request_id}] {mail.subject}")
        body = quote("Comentario (opcional): ") if word == "RECHAZAR" else ""
        return f"mailto:{approvals_mailbox}?subject={subj}&body={body}"

    esc = html.escape
    btn = ("display:inline-block;padding:10px 22px;margin-right:12px;border-radius:4px;"
           "color:#fff;text-decoration:none;font-weight:bold;")
    att = ", ".join(esc(a[0]) for a in mail.attachments) or "—"
    return f"""
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222">
  <p>Hola, se generó un correo del reporte <b>{esc(mail.report_name)}</b> que necesita tu aprobación.</p>
  <table style="border-collapse:collapse;font-size:13px">
    <tr><td style="padding:2px 10px 2px 0;color:#666">Solicitud</td><td><b>{request_id}</b></td></tr>
    <tr><td style="padding:2px 10px 2px 0;color:#666">Para</td><td>{esc(', '.join(mail.to))}</td></tr>
    <tr><td style="padding:2px 10px 2px 0;color:#666">CC</td><td>{esc(', '.join(mail.cc)) or '—'}</td></tr>
    <tr><td style="padding:2px 10px 2px 0;color:#666">Asunto</td><td>{esc(mail.subject)}</td></tr>
    <tr><td style="padding:2px 10px 2px 0;color:#666">Adjuntos</td><td>{att}</td></tr>
  </table>
  <p style="margin:18px 0">
    <a href="{mailto('APROBAR')}" style="{btn}background:#107c10">APROBAR</a>
    <a href="{mailto('RECHAZAR')}" style="{btn}background:#c50f1f">RECHAZAR</a>
  </p>
  <p><b>¿Quieres editarlo?</b> Abre el adjunto <i>propuesta.eml</i>: se abre como borrador
  con destinatarios y adjuntos. Tienes dos opciones:</p>
  <ol>
    <li>Edítalo y dale <b>Enviar</b> tú mismo (nosotros lo detectamos y cerramos la solicitud), o</li>
    <li>Responde a este correo con <b>APROBAR</b> y adjunta la versión editada
        (arrastra el borrador a tu respuesta). Se enviará la versión adjunta.</li>
  </ol>
  <p style="color:#666;font-size:12px">Si no respondes en {expires_hours} h la solicitud expira y no se envía nada.
  Abajo va la vista previa del cuerpo.</p>
  <hr>
  {mail.body_html}
</div>"""


# --------------------------------------------------------------------------- parseo
def _strip_reply_prefix(subject: str) -> str:
    return _REPLY_PREFIX.sub("", subject or "").strip()


def first_meaningful_line(body_text: str) -> str:
    """Primera línea con texto antes de la parte citada de la respuesta."""
    for line in (body_text or "").splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith(">") or re.match(r"^(De|From|Enviado|Sent)\s*:", s, re.IGNORECASE):
            break
        if re.match(r"^_{5,}|^-{5,}", s):
            break
        return s
    return ""


def parse_decision(subject: str, body_text: str) -> tuple[str | None, str | None, str]:
    """Devuelve (request_id, decision, comentario). decision ∈ {APPROVE, REJECT, None}."""
    m = REQ_ID_RE.search(subject or "") or REQ_ID_RE.search(body_text or "")
    req_id = m.group(0).upper() if m else None

    def classify(text: str) -> str | None:
        first = re.split(r"[\s\[\]:,.!-]+", text.strip().upper(), maxsplit=1)[0] if text.strip() else ""
        if first in REJECT_WORDS:
            return "REJECT"
        if first in APPROVE_WORDS:
            return "APPROVE"
        return None

    line = first_meaningful_line(body_text)
    clean_subject = _strip_reply_prefix(subject)
    decision = classify(clean_subject) or classify(line)
    comment = line
    if comment.lower().startswith("comentario (opcional):"):
        comment = comment.split(":", 1)[1].strip()
    return req_id, decision, comment


# --------------------------------------------------------------------------- flujo
def create_approval_requests(store: StateStore, client: MailClient, mails: Iterable[OutgoingEmail],
                             business_date: str, approvals_mailbox: str, expires_hours: int = 8) -> list[dict]:
    out = []
    for mail in mails:
        req_id = store.new_request_id()
        proposal = build_proposal_mime(mail, approvals_mailbox)
        body = build_request_html(req_id, mail, approvals_mailbox, expires_hours)
        client.send_html(
            to=[mail.oma_email],
            subject=f"[APROBACIÓN {req_id}] {mail.subject}",
            body_html=body,
            attachments=[("propuesta.eml", proposal, "message/rfc822")],
        )
        store.create_request(
            request_id=req_id, business_date=business_date, report_name=mail.report_name,
            oma_email=mail.oma_email, to_addrs=", ".join(mail.to), cc_addrs=", ".join(mail.cc),
            subject=mail.subject, proposal_mime=proposal,
        )
        out.append({"request_id": req_id, "oma": mail.oma_email, "subject": mail.subject, "status": "PENDING"})
    return out


def _addr_set(s: str) -> set[str]:
    return {a.strip().lower() for a in re.split(r"[;,]", s or "") if a.strip()}


def process_replies(store: StateStore, client: MailClient, approvals_mailbox: str,
                    delegates: dict[str, list[str]] | None = None) -> list[dict]:
    """Procesa respuestas en el buzón de aprobaciones. Devuelve un log (tabla para KNIME)."""
    delegates = {k.lower(): [d.lower() for d in v] for k, v in (delegates or {}).items()}
    log = []

    for msg in client.list_unprocessed():
        sender = (msg.sender or "").lower()
        req_id, decision, comment = parse_decision(msg.subject, msg.body_text)
        entry = {"message_id": msg.id, "from": sender, "subject": msg.subject,
                 "request_id": req_id, "decision": decision, "action": None, "detail": ""}

        if req_id is None:
            # ¿Es la copia BCC de una propuesta que el OMA editó y envió él mismo?
            match = _match_direct_send(store, msg, sender, delegates)
            if match:
                if store.set_status(match["request_id"], "EDITED_SENT", "Enviado directamente por el OMA"):
                    entry.update(request_id=match["request_id"], action="EDITED_SENT",
                                 detail="OMA editó y envió desde Outlook")
                else:
                    entry.update(request_id=match["request_id"], action="IGNORED", detail="Ya estaba cerrada")
            else:
                entry.update(action="IGNORED", detail="Sin ID de solicitud")
            client.mark_processed(msg.id)
            log.append(entry)
            continue

        row = store.get(req_id)
        if row is None:
            entry.update(action="IGNORED", detail="ID de solicitud desconocido")
        elif not _is_authorized(row["oma_email"], sender, delegates):
            entry.update(action="IGNORED", detail=f"Remitente no autorizado (OMA: {row['oma_email']})")
        elif row["status"] != "PENDING":
            entry.update(action="IGNORED", detail=f"Solicitud ya cerrada: {row['status']}")
        elif decision == "REJECT":
            store.set_status(req_id, "REJECTED", comment)
            entry.update(action="REJECTED", detail=comment)
        elif decision == "APPROVE":
            entry.update(**_send_approved(store, client, row, msg, approvals_mailbox))
        else:
            entry.update(action="IGNORED", detail="No se reconoció APROBAR/RECHAZAR")
            client.send_html(
                to=[sender], subject=f"No entendimos tu respuesta [{req_id}]",
                body_html=("<p>Responde con <b>APROBAR</b> o <b>RECHAZAR</b> en la primera línea "
                           f"o en el asunto, manteniendo el ID <b>{req_id}</b>.</p>"))
        client.mark_processed(msg.id)
        log.append(entry)
    return log


def _send_approved(store: StateStore, client: MailClient, row, msg: InboxMessage, approvals_mailbox: str) -> dict:
    req_id = row["request_id"]
    edited = client.get_email_attachments(msg.id) if msg.has_attachments else []
    mime, status = (edited[0], "EDITED_SENT") if edited else (row["proposal_mime"], "APPROVED_SENT")
    # Primero reservar (PENDING -> estado final) para nunca enviar dos veces.
    if not store.set_status(req_id, status, "Versión editada adjunta" if edited else "Aprobado"):
        return {"action": "IGNORED", "detail": "Otra ejecución ya la procesó"}
    try:
        client.send_mime(prepare_for_sending(mime, approvals_mailbox))
    except Exception as exc:  # noqa: BLE001 - se reporta en la tabla de KNIME
        store.force_status(req_id, "ERROR", f"Falló envío: {exc}")
        return {"action": "ERROR", "detail": str(exc)}
    return {"action": status, "detail": "Enviado"}


def _is_authorized(oma: str, sender: str, delegates: dict[str, list[str]]) -> bool:
    oma = (oma or "").lower()
    return sender == oma or sender in delegates.get(oma, [])


def _match_direct_send(store: StateStore, msg: InboxMessage, sender: str, delegates):
    subj = _strip_reply_prefix(msg.subject).lower()
    msg_to = {a.lower() for a in msg.to}
    for row in store.pending():
        if not _is_authorized(row["oma_email"], sender, delegates):
            continue
        if row["subject"].strip().lower() != subj:
            continue
        if msg_to and not (_addr_set(row["to_addrs"]) & msg_to):
            continue
        return row
    return None


def expire_old(store: StateStore, client: MailClient, max_age_hours: int, now: datetime | None = None) -> list[dict]:
    now = now or datetime.now(timezone.utc)
    out = []
    for row in store.pending():
        created = datetime.fromisoformat(row["created_ts"])
        if now - created >= timedelta(hours=max_age_hours):
            if store.set_status(row["request_id"], "EXPIRED", f"Sin respuesta en {max_age_hours} h"):
                client.send_html(
                    to=[row["oma_email"]],
                    subject=f"Solicitud expirada [{row['request_id']}] {row['subject']}",
                    body_html="<p>La solicitud expiró sin respuesta. El correo <b>no</b> se envió.</p>")
                out.append({"request_id": row["request_id"], "action": "EXPIRED"})
    return out
