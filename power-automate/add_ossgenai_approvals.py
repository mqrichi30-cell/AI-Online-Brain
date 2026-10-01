"""Insert an approval gate (Send email with options: Aprobar / Rechazar) before every
email the AWG Autoconsolidaciones flows send to ossgenai.im@pg.com.

Approve  -> the flow continues and the OSSGenAI email is sent exactly as shown.
Reject / timeout / approval error -> nothing is sent, the original email is put back
(unread) in the 'Marín, Cristhofer - AWG + Wakefern' folder and the run is cancelled.
"""
import glob
import json
import os
import re
import sys

ROOT = sys.argv[1]
WF = os.path.join(ROOT, "Workflows")

OSS = "ossgenai.im@pg.com"
OPT_APPROVE = "Aprobar"
OPT_REJECT = "Rechazar"
APPROVAL_TIMEOUT = "P1D"
CLIENT_FOLDER_NAME = "Marín, Cristhofer - AWG + Wakefern"
CLIENT_FOLDER_URI = (
    "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/childFolders?$filter=displayName%20eq%20"
    "%27Mar%C3%ADn%2C%20Cristhofer%20-%20AWG%20%2B%20Wakefern%27&$top=1"
)
SIG_ESCAPED = "@{replace(variables('SignatureHtml'),'\"','\\\"')}"
SIG_CLEAN = "@{variables('SignatureHtml')}"


def load(prefix):
    path = glob.glob(os.path.join(WF, prefix + "*.json"))[0]
    with open(path, encoding="utf-8") as fh:
        return path, json.load(fh)


def save(path, doc):
    with open(path, "w", encoding="utf-8", newline="\r\n") as fh:
        fh.write(json.dumps(doc, indent=2, ensure_ascii=False))


def find_parent(actions, name):
    """Return the actions dict that directly contains `name`."""
    if name in actions:
        return actions
    for v in actions.values():
        for sub in (v.get("actions"), (v.get("else") or {}).get("actions")):
            if sub:
                r = find_parent(sub, name)
                if r is not None:
                    return r
        for case in (v.get("cases") or {}).values():
            r = find_parent(case.get("actions", {}), name)
            if r is not None:
                return r
    return None


def template_to_object(template):
    """Turn a JSON body template with @{...} interpolations into a JSON object whose
    string values keep the interpolations (the signature no longer needs manual escaping)."""
    exprs = []

    def stash(m):
        exprs.append(m.group(0))
        return "__EXPR_%d__" % (len(exprs) - 1)

    tmp = template.replace(SIG_ESCAPED, "__SIG__")
    tmp = re.sub(r"@\{[^{}]*\}", stash, tmp)
    obj = json.loads(tmp)

    def restore(o):
        if isinstance(o, dict):
            return {k: restore(v) for k, v in o.items()}
        if isinstance(o, list):
            return [restore(v) for v in o]
        if isinstance(o, str):
            o = o.replace("__SIG__", SIG_CLEAN)
            for i, e in enumerate(exprs):
                o = o.replace("__EXPR_%d__" % i, e)
            assert "__EXPR_" not in o and "__SIG__" not in o
            return o
        return o

    return restore(obj)


def http(uri, method, body=None, immutable=True, auth=False, content_type=True):
    params = {"Uri": uri, "Method": method}
    if immutable:
        params["CustomHeader1"] = 'Prefer: IdType="ImmutableId"'
    if body is not None:
        params["Body"] = body
    if content_type:
        params["ContentType"] = "application/json"
    inputs = {
        "parameters": params,
        "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
            "operationId": "HttpRequest",
            "connectionName": "shared_office365",
        },
    }
    if auth:
        inputs["authentication"] = "@parameters('$authentication')"
    return {"type": "OpenApiConnection", "inputs": inputs}


def ra(**kw):
    return kw


def insert_gate(
    actions,
    *,
    p,  # action-name prefix, e.g. "AWG_-_CON_-_01_Aprobacion"
    run_after,  # runAfter for the first inserted action
    next_actions,  # actions that must now wait for the gate
    flow_label,
    step_label,
    thread_subject,  # expression
    reply_note,
    content_expr,  # expression rendering what will be sent
    original_label,
    mailbox_base,  # "v1.0/me" or "https://graph.microsoft.com/v1.0/users/x"
    original_id,  # expression for the original message id (immutable)
    client_folder_id,  # expression or None (looked up in the reject branch)
    auth,
):
    approval = p + "_Enviar_Solicitud"
    get_me = p + "_Get_Aprobador"
    gate = p + "_Guard_Aprobado"
    rej = p + "_Rechazo"

    actions[get_me] = dict(
        runAfter=run_after,
        **http(
            "https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName",
            "GET",
            immutable=False,
            auth=auth,
            content_type=False,
        ),
    )

    body_html = (
        "<p>Hola,</p>"
        "<p>El flujo <b>" + flow_label + "</b> está listo para enviar un correo a <b>" + OSS + "</b> "
        "y necesita tu aprobación.</p>"
        "<table cellpadding=\"4\" style=\"border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px\">"
        "<tr><td><b>Paso</b></td><td>" + step_label + "</td></tr>"
        "<tr><td><b>Para</b></td><td>" + OSS + "</td></tr>"
        "<tr><td><b>Asunto del hilo</b></td><td>@{coalesce(" + thread_subject + ", '')}</td></tr>"
        "<tr><td><b>Tipo de envío</b></td><td>" + reply_note + "</td></tr>"
        "<tr><td><b>Correo original</b></td><td>" + original_label + "</td></tr>"
        "</table>"
        "<p><b>" + OPT_APPROVE + "</b>: se envía el correo tal como aparece abajo y el flujo continúa.<br>"
        "<b>" + OPT_REJECT + "</b>: no se envía nada, no se notifica a nadie y el correo original vuelve como "
        "no leído a la carpeta <i>" + CLIENT_FOLDER_NAME + "</i>.<br>"
        "Si no respondes en 24 horas se trata como <b>" + OPT_REJECT + "</b>.</p>"
        "<hr><p><b>Contenido que se enviará a OSSGenAI:</b></p>"
        "<div style=\"border-left:4px solid #0078d4;padding:8px 12px;background:#f6f8fa\">@{" + content_expr + "}</div>"
    )
    send_inputs = {
        "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
            "operationId": "SendMailWithOptions",
            "connectionName": "shared_office365",
        },
        "parameters": {
            "optionsEmailSubscription/Message/To": "@{coalesce(body('" + get_me + "')?['mail'], body('" + get_me + "')?['userPrincipalName'])}",
            "optionsEmailSubscription/Message/Subject": "[AWG CON] Aprobación requerida - envío a OSSGenAI - " + step_label + " - @{coalesce(" + thread_subject + ", '')}",
            "optionsEmailSubscription/Message/Options": OPT_APPROVE + ", " + OPT_REJECT,
            "optionsEmailSubscription/Message/HeaderText": "Aprobación requerida: envío a OSSGenAI",
            "optionsEmailSubscription/Message/SelectionText": "¿Apruebas enviar este correo a OSSGenAI?",
            "optionsEmailSubscription/Message/Body": body_html,
            "optionsEmailSubscription/Message/Importance": "High",
            "optionsEmailSubscription/Message/UseOnlyHTMLMessage": True,
            "optionsEmailSubscription/Message/HideHTMLMessage": False,
            "optionsEmailSubscription/Message/ShowHTMLConfirmationDialog": False,
        },
    }
    if auth:
        send_inputs["authentication"] = "@parameters('$authentication')"
    actions[approval] = {
        "runAfter": {get_me: ["Succeeded"]},
        "limit": {"timeout": APPROVAL_TIMEOUT},
        "type": "OpenApiConnectionWebhook",
        "inputs": send_inputs,
    }

    # ---- reject branch -------------------------------------------------------
    rej_actions = {}
    folder_expr = client_folder_id
    first_ra = {}
    if folder_expr is None:
        rej_actions[rej + "_Get_Client_Folder"] = dict(
            runAfter={},
            **http(CLIENT_FOLDER_URI, "GET", immutable=False, auth=auth, content_type=False),
        )
        folder_expr = "first(body('" + rej + "_Get_Client_Folder')?['value'])?['id']"
        first_ra = {rej + "_Get_Client_Folder": ["Succeeded"]}

    msg_uri = mailbox_base + "/messages/@{encodeUriComponent(" + original_id + ")}"
    rej_actions[rej + "_Get_Original"] = dict(
        runAfter=first_ra,
        **http(msg_uri + "?$select=id,parentFolderId", "GET", auth=auth),
    )
    rej_actions[rej + "_Guard_Original_Fuera_De_Carpeta"] = {
        "runAfter": {rej + "_Get_Original": ["Succeeded"]},
        "type": "If",
        "expression": {
            "and": [
                {"not": {"equals": [
                    "@body('" + rej + "_Get_Original')?['parentFolderId']",
                    "@" + folder_expr,
                ]}}
            ]
        },
        "actions": {
            rej + "_Move_Original_To_Client_Folder": dict(
                **http(
                    msg_uri + "/move",
                    "POST",
                    "@concat('{\"destinationId\":\"', " + folder_expr + ", '\"}')",
                    auth=auth,
                ),
            )
        },
        "else": {"actions": {}},
    }
    rej_actions[rej + "_Mark_Original_Unread"] = dict(
        runAfter={rej + "_Guard_Original_Fuera_De_Carpeta": ["Succeeded", "Failed", "Skipped"]},
        **http(msg_uri, "PATCH", '{"isRead":false}', auth=auth),
    )
    rej_actions[rej + "_Terminate"] = {
        "runAfter": {rej + "_Mark_Original_Unread": ["Succeeded", "Failed", "Skipped"]},
        "type": "Terminate",
        "inputs": {"runStatus": "Cancelled"},
    }

    actions[gate] = {
        # Fail-safe: anything other than an explicit "Aprobar" (reject, timeout, error,
        # approval not sent) takes the reject branch, so nothing reaches OSSGenAI.
        "runAfter": {approval: ["Succeeded", "Failed", "TimedOut", "Skipped"]},
        "type": "If",
        "expression": {
            "and": [
                {"equals": ["@actions('" + approval + "')?['status']", "Succeeded"]},
                {"equals": ["@coalesce(actions('" + approval + "')?['outputs']?['body']?['SelectedOption'], '')", OPT_APPROVE]},
            ]
        },
        "actions": {
            p + "_Aprobado": {
                "runAfter": {},
                "type": "Compose",
                "inputs": "Aprobado por @{coalesce(actions('" + approval + "')?['outputs']?['body']?['UserEmailAddress'], '')}",
            }
        },
        "else": {"actions": rej_actions},
    }

    for n in next_actions:
        actions[n]["runAfter"] = {gate: ["Succeeded"]}


def payload_compose(actions, name, run_after, obj):
    actions[name] = {"runAfter": run_after, "type": "Compose", "inputs": obj}


# ============================================================================
# GENERAL - 00 : reply original customer email to OSSGenAI (classifier prompt)
# ============================================================================
path, doc = load("GENERAL-Classifier-00")
top = doc["properties"]["definition"]["actions"]
send_name = "GENERAL_-_00_HTTP_Reply_Original_to_OSSGenAI_SameThread"
acts = find_parent(top, send_name)
send = acts[send_name]
assert send.get("runAfter", {}) == {} and len(acts) == 1
payload = "GENERAL_-_00_Compose_OSSGenAI_Request"
payload_compose(acts, payload, {}, {
    "message": {
        "toRecipients": [{"emailAddress": {"address": OSS}}],
        "ccRecipients": [],
        "bccRecipients": [],
    },
    "comment": "@concat(outputs('GENERAL_-_00_Prompt_Guardrails'), outputs('GENERAL_-_00_Rule_Reply_Comment'), variables('SignatureHtml'))",
})
send["inputs"]["parameters"]["Body"] = "@outputs('" + payload + "')"
insert_gate(
    acts,
    p="GENERAL_-_00_Aprobacion",
    run_after={payload: ["Succeeded"]},
    next_actions=[send_name],
    flow_label="GENERAL - Classifier - 00 - Send Classify Prompt",
    step_label="Clasificación (prompt clasificador)",
    thread_subject="triggerOutputs()?['body/subject']",
    reply_note="Respuesta en el mismo hilo del correo del cliente (OSSGenAI recibe también el historial del hilo).",
    content_expr="outputs('" + payload + "')?['comment']",
    original_label="Correo del cliente que disparó el flujo (asunto arriba).",
    mailbox_base="v1.0/me",
    original_id="outputs('Compose_Original_Immutable_ID')",
    client_folder_id=None,
    auth=False,
)
save(path, doc)

# ============================================================================
# AWG - CON - 01 : reply current OSS message to OSSGenAI (consolidation prompt)
# ============================================================================
path, doc = load("AWG-Consolidation-01")
top = doc["properties"]["definition"]["actions"]
send_name = "AWG_-_CON_-_01_HTTP_Reply_Current_OSS_to_OSSGenAI_SameConversation"
acts = find_parent(top, send_name)
send = acts[send_name]
assert send.get("runAfter", {}) == {}
payload = "AWG_-_CON_-_01_Compose_OSSGenAI_Request"
payload_compose(acts, payload, {}, template_to_object(send["inputs"]["parameters"]["Body"]))
send["inputs"]["parameters"]["Body"] = "@outputs('" + payload + "')"
insert_gate(
    acts,
    p="AWG_-_CON_-_01_Aprobacion",
    run_after={payload: ["Succeeded"]},
    next_actions=[send_name],
    flow_label="AWG - Consolidation - 01 - Send Con Prompt",
    step_label="Consolidación 01 (extraer Truck # / PO #)",
    thread_subject="outputs('Compose_Original_Subject_From_Router')",
    reply_note="Respuesta en el mismo hilo (sobre la respuesta actual de OSSGenAI, incluye historial).",
    content_expr="outputs('" + payload + "')?['comment']",
    original_label="Correo original del cliente (carpeta Email Route 00).",
    mailbox_base="v1.0/me",
    original_id="outputs('Compose_Resolved_Original_Immutable_ID')",
    client_folder_id="outputs('AWG_-_CON_-_01_Target_Client_Folder_Id')",
    auth=False,
)
save(path, doc)

# ============================================================================
# AWG - CON - 03 : RDD alignment reply to OSSGenAI.
# The gate goes before the tracking update / MOC request so a rejection leaves
# nothing half done (no MOC sent, tracking untouched).
# ============================================================================
path, doc = load("AWG-Consolidation-03")
top = doc["properties"]["definition"]["actions"]
send_name = "AWG_-_CON_-_03_HTTP_Send_RDD_Reply_By_Immutable_ID"
acts = find_parent(top, send_name)
send = acts[send_name]
html = "AWG_-_CON_-_03_Compose_RDD_Reply_HTML"
assert acts[html]["runAfter"] == {"AWG_-_CON_-_03_Guard_Customer_Known": ["Succeeded", "Failed", "Skipped"]}
assert acts["Consolidation_Tracking"]["runAfter"] == {"DEBUG_All_POs": ["Succeeded"]}
acts[html]["runAfter"] = {"DEBUG_All_POs": ["Succeeded"]}
payload = "AWG_-_CON_-_03_Compose_OSSGenAI_Request"
payload_compose(acts, payload, {html: ["Succeeded"]}, {
    "message": {
        "toRecipients": [{"emailAddress": {"address": OSS}}],
        "body": {"contentType": "HTML", "content": "@{outputs('" + html + "')}"},
    }
})
send["inputs"]["parameters"]["Body"] = "@outputs('" + payload + "')"
send["runAfter"] = {"AWG_-_CON_-_03_Guard_Customer_Known": ["Succeeded", "Failed", "Skipped"]}
insert_gate(
    acts,
    p="AWG_-_CON_-_03_Aprobacion",
    run_after={payload: ["Succeeded"]},
    next_actions=["Consolidation_Tracking"],
    flow_label="AWG - Consolidation - 03 - Send MOC Form",
    step_label="Consolidación 03 (alineación de RDD)",
    thread_subject="triggerOutputs()?['body/subject']",
    reply_note="Respuesta en el mismo hilo del correo de ride-with. Si apruebas, también se envía el Consolidation Request (MOC) a nacsoshared como hasta ahora.",
    content_expr="outputs('" + html + "')",
    original_label="Correo de ride-with que disparó el flujo (asunto arriba).",
    mailbox_base="v1.0/me",
    original_id="outputs('AWG_-_CON_-_03_Compose_Current_Immutable_ID')",
    client_folder_id="outputs('AWG_-_CON_-_03_Target_Client_Folder_Id')",
    auth=False,
)
save(path, doc)

# ============================================================================
# AWG - CON - 04 : two replies to OSSGenAI (inside 10-day window / pending).
# Gates go before the SharePoint updates so a rejection leaves tracking untouched.
# ============================================================================
path, doc = load("AWG-Consolidation-04")
top = doc["properties"]["definition"]["actions"]
for send_name, first_name, suffix, step_label, note in [
    ("AWG_-_CON_-_04_Send_HTTP_OSS_Response", "Update_item_1", "Tabla",
     "Consolidación 04 (pedir tabla completa - dentro de 10 días)",
     "Respuesta en el mismo hilo de la respuesta de OSSGenAI (incluye historial)."),
    ("AWG_-_CON_-_04_Send_HTTP_OSS_Response_1", "Update_item", "10Dias",
     "Consolidación 04 (aviso pendiente 10 días)",
     "Respuesta en el mismo hilo de la respuesta de OSSGenAI (incluye historial). Si apruebas, también sale el aviso 'Pending 10 Days Out' al equipo como hasta ahora."),
]:
    acts = find_parent(top, send_name)
    send = acts[send_name]
    assert acts[first_name]["runAfter"] == {}
    payload = "AWG_-_CON_-_04_Compose_OSSGenAI_Request_" + suffix
    payload_compose(acts, payload, {}, template_to_object(send["inputs"]["parameters"]["Body"]))
    send["inputs"]["parameters"]["Body"] = "@outputs('" + payload + "')"
    insert_gate(
        acts,
        p="AWG_-_CON_-_04_Aprobacion_" + suffix,
        run_after={payload: ["Succeeded"]},
        next_actions=[first_name],
        flow_label="AWG - Consolidation - 04 - Query Con Status",
        step_label=step_label,
        thread_subject="triggerOutputs()?['body/subject']",
        reply_note=note,
        content_expr="outputs('" + payload + "')?['comment']",
        original_label="Respuesta de OSSGenAI que disparó el flujo (asunto arriba).",
        mailbox_base="https://graph.microsoft.com/v1.0/users/pgcustservw2.im@pg.com",
        original_id="outputs('AWG_-_CON_-_04_Compose_Current_Immutable_ID')",
        client_folder_id="outputs('AWG_-_CON_-_04_Target_Client_Folder_Id')",
        auth=True,
    )
save(path, doc)
print("ok")

# ============================================================================
# Trigger concurrency: GENERAL-00 and CON-03 had concurrency runs=1. With a run
# waiting on an approval (up to 24h), every new email queued behind it and
# Flow checker flagged "Trigger concurrency throttling". Remove it.
# ============================================================================
for prefix in ("GENERAL-Classifier-00", "AWG-Consolidation-03"):
    path, doc = load(prefix)
    for trig in doc["properties"]["definition"]["triggers"].values():
        trig.pop("runtimeConfiguration", None)
    save(path, doc)
print("concurrency removed")
