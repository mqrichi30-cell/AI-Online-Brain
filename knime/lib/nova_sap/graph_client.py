"""Cliente de Microsoft Graph (Outlook / Exchange Online) para el flujo de aprobaciones.

Requiere un App Registration en Entra ID con permisos de *aplicación*:
  - Mail.Send        (enviar desde el buzón emisor)
  - Mail.ReadWrite   (leer/marcar el buzón de aprobaciones)
Recomendado: restringir la app sólo a esos buzones con una Application Access Policy
(New-ApplicationAccessPolicy en Exchange Online).
"""
from __future__ import annotations

import base64
import re

import requests

from .approvals import InboxMessage

GRAPH = "https://graph.microsoft.com/v1.0"
PROCESSED_CATEGORY = "KNIME-Procesado"


class GraphMailClient:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str,
                 sender_mailbox: str, approvals_mailbox: str, timeout: int = 60):
        import msal  # import diferido para que los tests no requieran msal

        self._app = msal.ConfidentialClientApplication(
            client_id, authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret)
        self.sender = sender_mailbox
        self.approvals = approvals_mailbox
        self.timeout = timeout

    # ------------------------------------------------------------------ http
    def _token(self) -> str:
        res = self._app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in res:
            raise RuntimeError(f"No se pudo obtener token de Graph: {res.get('error_description', res)}")
        return res["access_token"]

    def _req(self, method: str, url: str, **kw) -> requests.Response:
        headers = kw.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._token()}"
        r = requests.request(method, url if url.startswith("http") else GRAPH + url,
                             headers=headers, timeout=self.timeout, **kw)
        if r.status_code >= 400:
            raise RuntimeError(f"Graph {method} {url} -> {r.status_code}: {r.text[:500]}")
        return r

    # ------------------------------------------------------------------ envío
    def send_html(self, to, subject, body_html, attachments=None) -> None:
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": a}} for a in to],
            "attachments": [
                {"@odata.type": "#microsoft.graph.fileAttachment", "name": n,
                 "contentType": ct, "contentBytes": base64.b64encode(b).decode()}
                for n, b, ct in (attachments or [])
            ],
        }
        self._req("POST", f"/users/{self.sender}/sendMail",
                  json={"message": message, "saveToSentItems": True})

    def send_mime(self, mime_bytes: bytes) -> None:
        # Graph acepta MIME en base64 con Content-Type text/plain (límite ~4 MB;
        # para adjuntos más grandes usar upload session).
        self._req("POST", f"/users/{self.sender}/sendMail",
                  data=base64.b64encode(mime_bytes), headers={"Content-Type": "text/plain"})

    # ------------------------------------------------------------------ lectura
    def list_unprocessed(self) -> list[InboxMessage]:
        url = (f"/users/{self.approvals}/mailFolders/inbox/messages?$top=50"
               f"&$filter=not(categories/any(c:c eq '{PROCESSED_CATEGORY}'))"
               "&$select=id,subject,from,toRecipients,body,hasAttachments,receivedDateTime"
               "&$orderby=receivedDateTime")
        out: list[InboxMessage] = []
        while url:
            data = self._req("GET", url, headers={"Prefer": 'outlook.body-content-type="text"'}).json()
            for m in data.get("value", []):
                out.append(InboxMessage(
                    id=m["id"],
                    sender=((m.get("from") or {}).get("emailAddress") or {}).get("address", ""),
                    subject=m.get("subject") or "",
                    body_text=(m.get("body") or {}).get("content", ""),
                    to=[r["emailAddress"]["address"] for r in m.get("toRecipients", [])],
                    has_attachments=bool(m.get("hasAttachments")),
                ))
            url = data.get("@odata.nextLink")
        return out

    def get_email_attachments(self, message_id: str) -> list[bytes]:
        """Devuelve los correos adjuntos como MIME:
        - itemAttachment (borrador arrastrado a la respuesta) -> /$value devuelve MIME
        - fileAttachment .eml -> contentBytes
        Los .msg sueltos no se soportan (formato binario de Outlook)."""
        base = f"/users/{self.approvals}/messages/{message_id}/attachments"
        atts = self._req("GET", base).json().get("value", [])
        out = []
        for a in atts:
            kind = a.get("@odata.type", "")
            if kind.endswith("itemAttachment"):
                out.append(self._req("GET", f"{base}/{a['id']}/$value").content)
            elif kind.endswith("fileAttachment") and re.search(r"\.eml$", a.get("name", ""), re.I):
                out.append(base64.b64decode(a["contentBytes"]))
        return out

    def mark_processed(self, message_id: str) -> None:
        self._req("PATCH", f"/users/{self.approvals}/messages/{message_id}",
                  json={"categories": [PROCESSED_CATEGORY], "isRead": True})
