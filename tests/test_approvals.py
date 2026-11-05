from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser

import pytest

from nova_sap.approvals import (
    InboxMessage, OutgoingEmail, build_proposal_mime, create_approval_requests, expire_old,
    parse_decision, prepare_for_sending, process_replies,
)
from nova_sap.state import StateStore

APPROVALS = "aprobaciones@empresa.com"
OMA = "oma@empresa.com"


class FakeMail:
    def __init__(self):
        self.sent_html, self.sent_mime, self.inbox, self.processed, self.attachments = [], [], [], [], {}

    def send_html(self, to, subject, body_html, attachments=None):
        self.sent_html.append({"to": to, "subject": subject, "body": body_html, "attachments": attachments or []})

    def send_mime(self, mime_bytes):
        self.sent_mime.append(BytesParser(policy=policy.SMTP).parsebytes(mime_bytes))

    def list_unprocessed(self):
        msgs = [m for m in self.inbox if m.id not in self.processed]
        return msgs

    def get_email_attachments(self, message_id):
        return self.attachments.get(message_id, [])

    def mark_processed(self, message_id):
        self.processed.append(message_id)


def sample_mail(subject="Fill rate semana 45"):
    return OutgoingEmail(report_name="Fill Rate", oma_email=OMA, to=["cliente@walmart.com"],
                         cc=["kam@empresa.com"], subject=subject, body_html="<p>Hola</p>",
                         attachments=[("reporte.xlsx", b"XLSXDATA", "application/vnd.ms-excel")])


@pytest.fixture
def env(tmp_path):
    store = StateStore(str(tmp_path / "s.sqlite"))
    mail = FakeMail()
    reqs = create_approval_requests(store, mail, [sample_mail()], "2026-11-05", APPROVALS)
    return store, mail, reqs[0]["request_id"]


def reply(mail, i, subject, body="", sender=OMA, has_att=False, to=None):
    mail.inbox.append(InboxMessage(id=f"m{i}", sender=sender, subject=subject, body_text=body,
                                   to=to or [APPROVALS], has_attachments=has_att))


def test_request_email_sent_with_editable_proposal(env):
    store, mail, rid = env
    req = mail.sent_html[0]
    assert req["to"] == [OMA] and rid in req["subject"]
    assert f"mailto:{APPROVALS}?subject=APROBAR%20%5B{rid}%5D" in req["body"]
    name, data, ctype = req["attachments"][0]
    assert name == "propuesta.eml" and ctype == "message/rfc822"
    proposal = BytesParser(policy=policy.SMTP).parsebytes(data)
    assert proposal["X-Unsent"] == "1" and proposal["Bcc"] == APPROVALS
    assert [p.get_filename() for p in proposal.iter_attachments()] == ["reporte.xlsx"]
    assert store.get(rid)["status"] == "PENDING"


def test_approve_sends_original_once(env):
    store, mail, rid = env
    reply(mail, 1, f"APROBAR [{rid}] Fill rate semana 45")
    log = process_replies(store, mail, APPROVALS)
    assert log[0]["action"] == "APPROVED_SENT"
    sent = mail.sent_mime[0]
    assert sent["To"] == "cliente@walmart.com" and sent["Cc"] == "kam@empresa.com"
    assert sent["X-Unsent"] is None and sent["Bcc"] is None
    # Segunda aprobación (doble clic) no reenvía
    reply(mail, 2, f"RE: APROBAR [{rid}]")
    log = process_replies(store, mail, APPROVALS)
    assert log[0]["action"] == "IGNORED" and len(mail.sent_mime) == 1
    assert store.get(rid)["status"] == "APPROVED_SENT"


def test_approve_with_edited_attachment_sends_edited(env):
    store, mail, rid = env
    edited = build_proposal_mime(sample_mail("Fill rate semana 45 (corregido)"), APPROVALS)
    reply(mail, 1, f"RE: [APROBACIÓN {rid}] Fill rate", body="Aprobado\n\nDe: robot", has_att=True)
    mail.attachments["m1"] = [edited]
    log = process_replies(store, mail, APPROVALS)
    assert log[0]["action"] == "EDITED_SENT"
    assert mail.sent_mime[0]["Subject"] == "Fill rate semana 45 (corregido)"


def test_reject_with_reason(env):
    store, mail, rid = env
    reply(mail, 1, f"RECHAZAR [{rid}] Fill rate", body="Comentario (opcional): cifras incorrectas")
    log = process_replies(store, mail, APPROVALS)
    assert log[0]["action"] == "REJECTED"
    row = store.get(rid)
    assert row["status"] == "REJECTED" and row["decision_note"] == "cifras incorrectas"
    assert mail.sent_mime == []


def test_unauthorized_sender_is_ignored(env):
    store, mail, rid = env
    reply(mail, 1, f"APROBAR [{rid}]", sender="otro@empresa.com")
    assert process_replies(store, mail, APPROVALS)[0]["action"] == "IGNORED"
    assert store.get(rid)["status"] == "PENDING"


def test_delegate_can_approve(env):
    store, mail, rid = env
    reply(mail, 1, f"APROBAR [{rid}]", sender="backup@empresa.com")
    log = process_replies(store, mail, APPROVALS, delegates={OMA: ["backup@empresa.com"]})
    assert log[0]["action"] == "APPROVED_SENT"


def test_direct_send_by_oma_detected_via_bcc(env):
    store, mail, rid = env
    reply(mail, 1, "Fill rate semana 45", body="texto editado", to=["cliente@walmart.com"])
    log = process_replies(store, mail, APPROVALS)
    assert log[0]["action"] == "EDITED_SENT" and log[0]["request_id"] == rid
    assert mail.sent_mime == []


def test_unrecognized_reply_asks_again(env):
    store, mail, rid = env
    reply(mail, 1, f"RE: [APROBACIÓN {rid}]", body="¿qué es esto?")
    log = process_replies(store, mail, APPROVALS)
    assert log[0]["action"] == "IGNORED"
    assert "No entendimos" in mail.sent_html[-1]["subject"]
    assert store.get(rid)["status"] == "PENDING"


def test_expire(env):
    store, mail, rid = env
    out = expire_old(store, mail, 8, now=datetime.now(timezone.utc) + timedelta(hours=9))
    assert out == [{"request_id": rid, "action": "EXPIRED"}]
    assert store.get(rid)["status"] == "EXPIRED"


def test_send_failure_marks_error(env):
    store, mail, rid = env

    def boom(_):
        raise RuntimeError("Graph 503")
    mail.send_mime = boom
    reply(mail, 1, f"APROBAR [{rid}]")
    assert process_replies(store, mail, APPROVALS)[0]["action"] == "ERROR"
    assert store.get(rid)["status"] == "ERROR"


@pytest.mark.parametrize("subject,body,decision", [
    ("APROBAR [REQ-ABCDEF12] x", "", "APPROVE"),
    ("RE: [APROBACIÓN REQ-ABCDEF12] x", "Aprobado\n> citado", "APPROVE"),
    ("RE: [APROBACIÓN REQ-ABCDEF12] x", "ok, pero cambia la tabla", None),
    ("RV: RE: [APROBACIÓN REQ-ABCDEF12] x", "Rechazado, falta info", "REJECT"),
    ("RE: [APROBACIÓN REQ-ABCDEF12] x", "\n\nDe: alguien\nAPROBAR", None),
])
def test_parse_decision(subject, body, decision):
    rid, d, _ = parse_decision(subject, body)
    assert rid == "REQ-ABCDEF12" and d == decision


def test_prepare_for_sending_requires_to():
    with pytest.raises(ValueError):
        prepare_for_sending(b"Subject: x\r\n\r\nhola")
