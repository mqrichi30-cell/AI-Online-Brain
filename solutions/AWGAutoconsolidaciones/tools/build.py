"""Builds the two AWGAutoconsolidaciones variants:
  outlook : Try/Catch on every flow + approvals before every OSSGenAI send (Outlook sends kept)
  smtp    : same, but every email send uses the SMTP connector and approvals use the Approvals connector
usage: python3 -I build.py <extracted_solution_dir> <trycatch_definition.json> <out_dir> <variant>
"""
import copy, json, os, re, shutil, sys

SRC, TC, OUT, VARIANT = sys.argv[1:5]
assert VARIANT in ('outlook', 'smtp')
SMTP = VARIANT == 'smtp'
NEW_VERSION = '1.0.0.38'
OSS = 'ossgenai.im@pg.com'
SHARED_MAILBOX = 'pgcustservw2.im@pg.com'

SMTP_REF = 'awg_sharedsmtp_5f1c2'
APPROVALS_REF = 'awg_sharedapprovals_8b3e4'
CONVERSION_REF = 'awg_sharedconversionservice_360cd'

O365_HOST = {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
             "operationId": "HttpRequest", "connectionName": "shared_office365"}

# ---------------------------------------------------------------- catch template (copied verbatim)
tc = json.load(open(TC, encoding='utf-8-sig'))
tc_actions = tc['properties']['definition']['actions']
CATCH = tc_actions['Scope_-_Catch']
TRY_NAME = 'Scope_-_Try'
CATCH_NAME = 'Scope_-_Catch'


# ---------------------------------------------------------------- helpers
def iter_containers(actions):
    """yield every actions-dict (top level, scopes, if/else branches, foreach, switch)."""
    yield actions
    for v in actions.values():
        if 'actions' in v:
            yield from iter_containers(v['actions'])
        if 'else' in v and 'actions' in v['else']:
            yield from iter_containers(v['else']['actions'])
        for c in v.get('cases', {}).values():
            yield from iter_containers(c.get('actions', {}))
        if 'default' in v:
            yield from iter_containers(v['default'].get('actions', {}))


def all_names(actions):
    names = []
    for c in iter_containers(actions):
        names.extend(c.keys())
    return names


def find(actions, name):
    for c in iter_containers(actions):
        if name in c:
            return c
    raise KeyError(name)


def successors(container, name):
    return [k for k, v in container.items() if name in (v.get('runAfter') or {})]


def insert_before(container, target, new_actions):
    """Chain new_actions (ordered dict) in front of target: they inherit target's runAfter,
    target then runs after the last new action."""
    first = True
    prev = None
    for k, v in new_actions.items():
        assert k not in container, k
        if first:
            v['runAfter'] = container[target].get('runAfter') or {}
            first = False
        else:
            v.setdefault('runAfter', {prev: ['Succeeded']})
        container[k] = v
        prev = k
    container[target]['runAfter'] = {prev: ['Succeeded']}


def o365_get(uri):
    return {"type": "OpenApiConnection",
            "inputs": {"parameters": {"Uri": uri, "Method": "GET",
                                      "CustomHeader1": "Prefer: IdType=\"ImmutableId\"",
                                      "ContentType": "application/json"},
                       "host": dict(O365_HOST)}}


def smtp_send(to, subject, body, importance='Normal', cc=None, attachments=None, sender=None):
    p = {"emailMessage/From": sender or "@variables('SmtpFrom')",
         "emailMessage/To": to,
         "emailMessage/Subject": subject,
         "emailMessage/Body": body,
         "emailMessage/Importance": importance,
         "emailMessage/IsHtml": True}
    if cc:
        p["emailMessage/CC"] = cc
    if attachments:
        p["emailMessage/Attachments"] = attachments
    return {"type": "OpenApiConnection",
            "inputs": {"parameters": p,
                       "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_smtp",
                                "operationId": "SendEmailV3", "connectionName": "shared_smtp"}}}


def approval_request_outlook(name_get, title_subject, header_html):
    return {"limit": {"timeout": "P1D"}, "type": "OpenApiConnectionWebhook",
            "inputs": {"host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
                                "operationId": "SendMailWithOptions", "connectionName": "shared_office365"},
                       "parameters": {
                           "optionsEmailSubscription/Message/To": "@{coalesce(body('%s')?['mail'], body('%s')?['userPrincipalName'])}" % (name_get, name_get),
                           "optionsEmailSubscription/Message/Subject": title_subject,
                           "optionsEmailSubscription/Message/Options": "Aprobar, Rechazar",
                           "optionsEmailSubscription/Message/HeaderText": "Aprobación requerida: envío a OSSGenAI",
                           "optionsEmailSubscription/Message/SelectionText": "¿Apruebas enviar este correo a OSSGenAI?",
                           "optionsEmailSubscription/Message/Body": header_html,
                           "optionsEmailSubscription/Message/Importance": "High",
                           "optionsEmailSubscription/Message/UseOnlyHTMLMessage": True,
                           "optionsEmailSubscription/Message/HideHTMLMessage": False,
                           "optionsEmailSubscription/Message/ShowHTMLConfirmationDialog": False}}}


def approval_request_approvals(name_get, title, details_expr):
    """Approvals connector: actionable Approve / Reject e-mail (not sent through Outlook)."""
    return {"limit": {"timeout": "P1D"}, "type": "OpenApiConnectionWebhook",
            "inputs": {"host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_approvals",
                                "operationId": "StartAndWaitForAnApproval", "connectionName": "shared_approvals"},
                       "parameters": {
                           "approvalType": "Basic",
                           "WebhookApprovalCreationInput/title": title,
                           "WebhookApprovalCreationInput/assignedTo": "@{coalesce(body('%s')?['mail'], body('%s')?['userPrincipalName'])}" % (name_get, name_get),
                           "WebhookApprovalCreationInput/details": details_expr,
                           "WebhookApprovalCreationInput/enableNotifications": True,
                           "WebhookApprovalCreationInput/enableReassignment": False}}}


def html_to_text(html_expr):
    return {"type": "OpenApiConnection",
            "inputs": {"parameters": {"Content": html_expr},
                       "host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_conversionservice",
                                "operationId": "HtmlToText", "connectionName": "shared_conversionservice"}}}


def md_details(text_action):
    # Approvals "details" is Markdown: a blank line is needed to keep each line break.
    nl = "decodeUriComponent('%0A')"
    return "@{replace(coalesce(body('%s'), ''), %s, concat(%s, %s))}" % (text_action, nl, nl, nl)


# ---------------------------------------------------------------- SMTP: existing approvals -> Approvals connector
def convert_existing_approvals(actions, used):
    for c in list(iter_containers(actions)):
        for k in list(c.keys()):
            v = c[k]
            if v.get('type') == 'OpenApiConnectionWebhook' and v['inputs']['host']['operationId'] == 'SendMailWithOptions':
                p = v['inputs']['parameters']
                to = p['optionsEmailSubscription/Message/To']
                get_name = re.search(r"body\('([^']+)'\)", to).group(1)
                prefix = k[:-len('Enviar_Solicitud')]
                html_name = prefix + 'Detalle_HTML'
                txt_name = prefix + 'Detalle_Texto'
                assert html_name not in used and txt_name not in used
                c[html_name] = {"runAfter": v['runAfter'], "type": "Compose",
                                "inputs": "<p><b>%s</b></p>%s" % (p['optionsEmailSubscription/Message/Subject'],
                                                                  p['optionsEmailSubscription/Message/Body'])}
                c[txt_name] = dict(html_to_text("@{outputs('%s')}" % html_name), runAfter={html_name: ['Succeeded']})
                title = p['optionsEmailSubscription/Message/Subject'].split(' - @{')[0]
                new = approval_request_approvals(get_name, title, md_details(txt_name))
                new['runAfter'] = {txt_name: ['Succeeded']}
                c[k] = new
                used.update([html_name, txt_name])
    # guards: SelectedOption 'Aprobar'  ->  outcome 'Approve'
    for c in iter_containers(actions):
        for k, v in c.items():
            if k.endswith('Aprobacion_Guard_Aprobado') or re.search(r'Aprobacion_\w+_Guard_Aprobado$', k):
                s = json.dumps(v, ensure_ascii=False)
                s = s.replace("?['outputs']?['body']?['SelectedOption'], '')\", \"Aprobar\"",
                              "?['outputs']?['body']?['outcome'], '')\", \"Approve\"")
                s = s.replace("?['outputs']?['body']?['UserEmailAddress']",
                              "?['outputs']?['body']?['responses']?[0]?['responder']?['email']")
                c[k] = json.loads(s)


# ---------------------------------------------------------------- SMTP: convert sends
def convert_sends(actions, flow_tag, used, needs_from):
    counter = [0]

    def nm(suffix):
        counter[0] += 1
        n = '%s_SMTP_%d_%s' % (flow_tag, counter[0], suffix)
        assert n not in used and len(n) <= 80, n
        used.add(n)
        return n

    for c in list(iter_containers(actions)):
        for k in list(c.keys()):
            if k not in c:
                continue
            v = c[k]
            if v.get('type') != 'OpenApiConnection':
                continue
            host = v['inputs']['host']
            if 'shared_office365' not in host['apiId']:
                continue
            op = host['operationId']
            p = v['inputs']['parameters']
            if op == 'SendEmailV2':
                new = smtp_send(p['emailMessage/To'], p['emailMessage/Subject'], p['emailMessage/Body'],
                                p.get('emailMessage/Importance', 'Normal'), cc=p.get('emailMessage/Cc'))
                new['runAfter'] = v.get('runAfter')
                if v.get('metadata'):
                    new['metadata'] = v['metadata']
                c[k] = new
                needs_from.add(True)
            elif op == 'HttpRequest' and p.get('Method') == 'POST' and re.search(r"/reply$", p['Uri']):
                convert_reply(c, k, nm, needs_from)
            elif op == 'HttpRequest' and p.get('Method') == 'POST' and p['Uri'] == 'v1.0/me/messages':
                # flow 03: draft with attachment + /send  ->  one SMTP send with the attachment
                assert k == 'AWG_-_CON_-_03_HTTP_Send_Consolidation_Request', k
                sender = c['AWG_-_CON_-_03_HTTP_Send_Created_Consolidation_Message']
                assert successors(c, 'AWG_-_CON_-_03_HTTP_Send_Created_Consolidation_Message') == []
                del c['AWG_-_CON_-_03_HTTP_Send_Created_Consolidation_Message']
                new = smtp_send("nacsoshared.im@pg.com",
                                "@{outputs('AWG_-_CON_-_03_Compose_Customer_Final')} Consolidation Request",
                                "Please process attached consolidation request.<br><br><br><br>@{variables('SignatureHtml')}",
                                attachments=[{"FileName": "MOC.xlsx",
                                              "ContentData": "@outputs('AWG_-_CON_-_03_Compose_MOC_Attachment_Base64')"}])
                new['runAfter'] = v.get('runAfter')
                c[k] = new
                needs_from.add(True)
            elif op in ('SendMailWithOptions', 'SendEmail', 'ForwardEmail_V2', 'ReplyToV3'):
                raise RuntimeError('unhandled send ' + k)
            elif op == 'HttpRequest' and p.get('Method') == 'POST' and not re.search(r"/move['\")}]*$", p['Uri']):
                raise RuntimeError('unhandled POST ' + k + ' ' + p['Uri'])


def convert_reply(c, k, nm, needs_from):
    """Graph /reply  ->  read original (GET) + build recipients/subject/history + SMTP send."""
    v = c[k]
    p = v['inputs']['parameters']
    uri = p['Uri']
    get_uri = uri[:-len('/reply')] + "?$select=subject,from,replyTo,toRecipients,ccRecipients,sentDateTime,body"
    body_expr = p['Body']
    assert body_expr.startswith('@') and not body_expr.startswith('@{'), body_expr
    shared = SHARED_MAILBOX in uri
    sender = SHARED_MAILBOX if shared else None
    if not shared:
        needs_from.add(True)
    n_req = nm('Reply_Request')
    n_get = nm('Get_Original')
    n_to = nm('Select_To')
    n_cc = nm('Select_Cc')
    n_oto = nm('Select_Original_To')
    n_html = nm('Compose_Body')
    empty = "json('[]')"
    to_src = ("@if(greater(length(coalesce(outputs('{r}')?['message']?['toRecipients'], {e})), 0), "
              "outputs('{r}')?['message']?['toRecipients'], "
              "if(greater(length(coalesce(body('{g}')?['replyTo'], {e})), 0), body('{g}')?['replyTo'], "
              "createArray(body('{g}')?['from'])))").format(r=n_req, g=n_get, e=empty)
    html = ("@concat(coalesce(outputs('{r}')?['comment'], outputs('{r}')?['message']?['body']?['content'], ''), "
            "'<br><br><hr style=\"border:none;border-top:1px solid #e1e1e1\">"
            "<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt\"><b>From:</b> ', "
            "coalesce(body('{g}')?['from']?['emailAddress']?['name'], ''), ' &lt;', "
            "coalesce(body('{g}')?['from']?['emailAddress']?['address'], ''), '&gt;<br><b>Sent:</b> ', "
            "coalesce(body('{g}')?['sentDateTime'], ''), '<br><b>To:</b> ', join(body('{ot}'), '; '), "
            "'<br><b>Subject:</b> ', coalesce(body('{g}')?['subject'], ''), '</div><br>', "
            "coalesce(body('{g}')?['body']?['content'], ''))").format(r=n_req, g=n_get, ot=n_oto)
    subject = ("@{{if(startsWith(toLower(trim(coalesce(body('{g}')?['subject'], ''))), 're:'), "
               "body('{g}')?['subject'], concat('RE: ', coalesce(body('{g}')?['subject'], '')))}}").format(g=n_get)
    pre = {
        n_req: {"type": "Compose", "inputs": "@json(string(%s))" % body_expr[1:]},
        n_get: dict(o365_get(get_uri)),
        n_to: {"type": "Select", "inputs": {"from": to_src, "select": "@item()?['emailAddress']?['address']"}},
        n_cc: {"type": "Select", "inputs": {"from": "@coalesce(outputs('%s')?['message']?['ccRecipients'], %s)" % (n_req, empty),
                                            "select": "@item()?['emailAddress']?['address']"}},
        n_oto: {"type": "Select", "inputs": {"from": "@coalesce(body('%s')?['toRecipients'], %s)" % (n_get, empty),
                                             "select": "@item()?['emailAddress']?['address']"}},
        n_html: {"type": "Compose", "inputs": html},
    }
    new = smtp_send("@{join(body('%s'), ';')}" % n_to, subject, "@{outputs('%s')}" % n_html,
                    cc="@{join(body('%s'), ';')}" % n_cc, sender=sender)
    new['runAfter'] = v.get('runAfter')
    if v.get('metadata'):
        new['metadata'] = v['metadata']
    c[k] = new
    insert_before(c, k, pre)


# ---------------------------------------------------------------- flow 05: conditional approval
def add_flow05_approvals(actions, used):
    """Flow 05 replies to the customer's original e-mail, but when the tracking row has no
    OriginalMessageId it falls back to the current e-mail, which comes from OSSGenAI.
    Ask for approval whenever the reply would reach OSSGenAI."""
    cond3 = find(actions, 'AWG_-_CON_-_05_HTTP_Reply_Customer_Same_Thread')
    specs = [
        (cond3, 'AWG_-_CON_-_05_HTTP_Reply_Customer_Same_Thread', 'AWG_-_CON_-_05_Aprobacion_Cliente_',
         'AWG_-_CON_-_05_Compose_Customer_Reply_HTML',
         "string(body('AWG_-_CON_-_05_Select_Cc_Objects'))", 'Confirmación al cliente'),
    ]
    else3 = find(actions, 'AWG_-_CON_-_05_HTTP_Reply_NoCollective_Same_Thread')
    specs.append((else3, 'AWG_-_CON_-_05_HTTP_Reply_NoCollective_Same_Thread', 'AWG_-_CON_-_05_Aprobacion_NoCol_',
                  'AWG_-_CON_-_05_Compose_NoCollective_HTML', "''", 'Aviso sin collective'))
    for c, target, pre, html_action, cc_expr, label in specs:
        n_dest = pre + 'Get_Destino'
        n_check = pre + 'Requiere'
        n_get = pre + 'Get_Aprobador'
        n_req = pre + 'Enviar_Solicitud'
        n_guard = pre + 'Guard_Aprobado'
        for n in (n_dest, n_check, n_get, n_req, n_guard):
            assert n not in used and len(n) <= 80
        dest = o365_get("v1.0/me/messages/@{encodeUriComponent(outputs('AWG_-_CON_-_05_Compose_Reply_Target_Id'))}?$select=from,replyTo,subject")
        para = ("@{coalesce(first(coalesce(body('%s')?['replyTo'], json('[]')))?['emailAddress']?['address'], "
                "body('%s')?['from']?['emailAddress']?['address'], '')}" % (n_dest, n_dest))
        header = ("<p>Hola,</p><p>El flujo <b>AWG - Consolidation - 05 - Confirm Con Client</b> va a responder un correo "
                  "cuyo destinatario incluye a <b>%s</b> y necesita tu aprobación.</p>"
                  "<table cellpadding=\"4\" style=\"border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px\">"
                  "<tr><td><b>Paso</b></td><td>Consolidación 05 (%s)</td></tr>"
                  "<tr><td><b>Remitente del correo que se responde</b></td><td>%s</td></tr>"
                  "<tr><td><b>Asunto del hilo</b></td><td>@{coalesce(body('%s')?['subject'], '')}</td></tr></table>"
                  "<p><b>Aprobar</b>: se envía el correo tal como aparece abajo y el flujo continúa.<br>"
                  "<b>Rechazar</b>: no se envía nada y el correo de OSSGenAI queda como no leído.<br>"
                  "Si no respondes en 24 horas se trata como <b>Rechazar</b>.</p><hr>"
                  "<p><b>Contenido que se enviará:</b></p>"
                  "<div style=\"border-left:4px solid #0078d4;padding:8px 12px;background:#f6f8fa\">@{outputs('%s')}</div>"
                  % (OSS, label, para, n_dest, html_action))
        subject = "[AWG CON] Aprobación requerida - envío a OSSGenAI - Consolidación 05 (%s) - @{coalesce(body('%s')?['subject'], '')}" % (label, n_dest)
        inner = {n_get: dict(o365_get("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName"), runAfter={})}
        inner[n_get]['inputs']['parameters'].pop('CustomHeader1')
        inner[n_get]['inputs']['parameters'].pop('ContentType')
        if SMTP:
            n_html, n_txt = pre + 'Detalle_HTML', pre + 'Detalle_Texto'
            inner[n_html] = {"runAfter": {n_get: ['Succeeded']}, "type": "Compose",
                             "inputs": "<p><b>%s</b></p>%s" % (subject, header)}
            inner[n_txt] = dict(html_to_text("@{outputs('%s')}" % n_html), runAfter={n_html: ['Succeeded']})
            inner[n_req] = approval_request_approvals(n_get, subject.split(' - @{')[0], md_details(n_txt))
            inner[n_req]['runAfter'] = {n_txt: ['Succeeded']}
            ok_expr = [{"equals": ["@actions('%s')?['status']" % n_req, "Succeeded"]},
                       {"equals": ["@coalesce(actions('%s')?['outputs']?['body']?['outcome'], '')" % n_req, "Approve"]}]
        else:
            inner[n_req] = approval_request_outlook(n_get, subject, header)
            inner[n_req]['runAfter'] = {n_get: ['Succeeded']}
            ok_expr = [{"equals": ["@actions('%s')?['status']" % n_req, "Succeeded"]},
                       {"equals": ["@coalesce(actions('%s')?['outputs']?['body']?['SelectedOption'], '')" % n_req, "Aprobar"]}]
        inner[n_guard] = {
            "runAfter": {n_req: ['Succeeded', 'Failed', 'TimedOut', 'Skipped']}, "type": "If",
            "expression": {"and": ok_expr},
            "actions": {pre + 'Aprobado': {"runAfter": {}, "type": "Compose", "inputs": "Aprobado"}},
            "else": {"actions": {
                pre + 'Rechazo_Mark_Unread': dict(o365_get("v1.0/me/messages/@{encodeUriComponent(outputs('AWG_-_CON_-_05_Compose_Current_Immutable_ID'))}"), runAfter={}),
                pre + 'Rechazo_Terminate': {"runAfter": {pre + 'Rechazo_Mark_Unread': ['Succeeded', 'Failed', 'TimedOut', 'Skipped']},
                                            "type": "Terminate", "inputs": {"runStatus": "Cancelled"}}}}}
        mu = inner[n_guard]['else']['actions'][pre + 'Rechazo_Mark_Unread']['inputs']['parameters']
        mu['Method'] = 'PATCH'
        mu['Body'] = '{"isRead":false}'
        check = {"type": "If",
                 "expression": {"or": [
                     {"contains": ["@toLower(concat(string(body('%s')?['from']), string(body('%s')?['replyTo']), %s))" % (n_dest, n_dest, cc_expr), OSS]}]},
                 "actions": inner, "else": {"actions": {}}}
        insert_before(c, target, {n_dest: dest, n_check: check})
        used.update([n_dest, n_check, n_get, n_req, n_guard] + list(inner.keys()) + list(inner[n_guard]['else']['actions'].keys()))


# ---------------------------------------------------------------- flow 00: 24h guard after the business-hours wait
def fix_flow00_only_new_email(actions):
    """Guard_Only_New_Email skips e-mails received more than 24h before *now*, but it runs after
    'Esperar horario laboral', which can hold the run up to ~62h (Friday after 3pm -> Monday 5am).
    Measure the 24h against the moment the trigger fired instead, so the weekend wait does not
    turn every pending e-mail into an 'old' one that is silently skipped."""
    c = find(actions, 'GENERAL_-_00_Guard_Only_New_Email')
    g = c['GENERAL_-_00_Guard_Only_New_Email']
    s = json.dumps(g['expression'])
    old = "@ticks(addMinutes(utcNow(), -1440))"
    assert s.count(old) == 1, s
    g['expression'] = json.loads(s.replace(old, "@ticks(addMinutes(trigger()?['startTime'], -1440))"))


# ---------------------------------------------------------------- try / catch
def wrap_try_catch(df, extra_vars):
    acts = df['actions']
    var_names = [k for k, v in acts.items() if v['type'] == 'InitializeVariable']
    # order variables by their original position in the run chain (topological)
    order, seen = [], set()

    def visit(n):
        if n in seen:
            return
        seen.add(n)
        for p in (acts[n].get('runAfter') or {}):
            visit(p)
        order.append(n)
    for n in acts:
        visit(n)
    var_order = [n for n in order if n in var_names]

    def real_preds(n, statuses=None):
        """predecessors of n with variable actions replaced by their own (non-variable) predecessors."""
        out = {}
        for p, st in (acts[n].get('runAfter') or {}).items():
            if p in var_names:
                out.update(real_preds(p, st))
            else:
                out[p] = st
        return out

    new_ra = {k: real_preds(k) for k, v in acts.items() if k not in var_names}

    top = {}
    prev = None
    for n in var_order:
        v = copy.deepcopy(acts[n])
        v['runAfter'] = {prev: ['Succeeded']} if prev else {}
        top[n] = v
        prev = n
    for name, value_def in extra_vars:
        top[name] = {"runAfter": {prev: ['Succeeded']} if prev else {}, "type": "InitializeVariable",
                     "inputs": {"variables": [value_def]}}
        prev = name

    inner = {}
    for k, v in acts.items():
        if k in var_names:
            continue
        v = copy.deepcopy(v)
        v['runAfter'] = new_ra[k]
        inner[k] = v

    collisions = (set(all_names(inner)) | set(top)) & (set(all_names(CATCH['actions'])) | {TRY_NAME, CATCH_NAME})
    assert not collisions, collisions

    catch = copy.deepcopy(CATCH)
    assert catch['runAfter'] == {"Scope_-_Try": ["TimedOut", "Failed"]}
    top[TRY_NAME] = {"actions": inner, "runAfter": {prev: ['Succeeded']} if prev else {}, "type": "Scope"}
    top[CATCH_NAME] = catch
    df['actions'] = top
    return inner


def add_smtp_from(inner, flow_tag, used):
    """SmtpFrom = mailbox of the Outlook connection (the same mailbox Outlook used to send from,
    so OSSGenAI answers keep arriving to the folders the triggers watch)."""
    n_get = flow_tag + '_SMTP_Get_Remitente'
    n_set = flow_tag + '_SMTP_Set_Remitente'
    assert n_get not in used and n_set not in used
    roots = [k for k, v in inner.items() if not v.get('runAfter')]
    g = o365_get("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName")
    g['inputs']['parameters'].pop('CustomHeader1')
    g['runAfter'] = {}
    s = {"runAfter": {n_get: ['Succeeded']}, "type": "SetVariable",
         "inputs": {"name": "SmtpFrom", "value": "@{coalesce(body('%s')?['mail'], body('%s')?['userPrincipalName'])}" % (n_get, n_get)}}
    for r in roots:
        inner[r]['runAfter'] = {n_set: ['Succeeded']}
    new_inner = {n_get: g, n_set: s}
    new_inner.update(inner)
    inner.clear()
    inner.update(new_inner)


# ---------------------------------------------------------------- per flow
def tag_of(fname):
    m = re.match(r'AWG-Consolidation-(\d\d)', fname)
    return 'AWG_-_CON_-_%s' % m.group(1) if m else 'GENERAL_-_00'


def process_flow(path):
    raw = open(path, encoding='utf-8-sig').read()
    d = json.loads(raw)
    df = d['properties']['definition']
    refs = d['properties']['connectionReferences']
    fname = os.path.basename(path)
    tag = tag_of(fname)
    used = set(all_names(df['actions']))

    if fname.startswith('GENERAL-Classifier-00'):
        fix_flow00_only_new_email(df['actions'])

    if fname.startswith('AWG-Consolidation-05'):
        add_flow05_approvals(df['actions'], used)

    needs_from = set()
    if SMTP:
        convert_existing_approvals(df['actions'], used)
        convert_sends(df['actions'], tag, used, needs_from)

    extra_vars = []
    if needs_from:
        extra_vars.append(('Initialize_variable_-_SmtpFrom', {"name": "SmtpFrom", "type": "string", "value": ""}))
    inner = wrap_try_catch(df, extra_vars)
    if needs_from:
        add_smtp_from(inner, tag, used)

    s = json.dumps(df)
    if '"shared_smtp"' in s:
        refs['shared_smtp'] = {"runtimeSource": "embedded", "connection": {"connectionReferenceLogicalName": SMTP_REF},
                               "api": {"name": "shared_smtp"}}
    if '"shared_approvals"' in s:
        refs['shared_approvals'] = {"runtimeSource": "embedded", "connection": {"connectionReferenceLogicalName": APPROVALS_REF},
                                    "api": {"name": "shared_approvals"}}
    if '"shared_conversionservice"' in s and 'shared_conversionservice' not in refs:
        refs['shared_conversionservice'] = {"runtimeSource": "embedded",
                                            "connection": {"connectionReferenceLogicalName": CONVERSION_REF},
                                            "api": {"name": "shared_conversionservice"}}
    if SMTP:
        # guard: no Outlook action may send mail in the SMTP variant
        for c in iter_containers(df['actions']):
            for k, v in c.items():
                h = v.get('inputs', {}).get('host', {}) if isinstance(v.get('inputs'), dict) else {}
                if 'shared_office365' in h.get('apiId', ''):
                    op = h['operationId']
                    pp = v['inputs'].get('parameters', {})
                    assert op in ('HttpRequest', 'OnNewEmailV3', 'SharedMailboxOnNewEmailV2', 'MoveV2'), (k, op)
                    if op == 'HttpRequest':
                        assert not re.search(r"/(reply|replyAll|createReply|forward|createForward|send|sendMail)['\")}]*$", pp['Uri']), (k, pp['Uri'])
                        assert not (pp['Method'] == 'POST' and pp['Uri'].rstrip('/').endswith('messages')), k
    out = json.dumps(d, indent=2, ensure_ascii=False).replace('\n', '\r\n')
    open(path, 'w', encoding='utf-8', newline='').write(out)
    return fname, refs


# ---------------------------------------------------------------- main
if os.path.exists(OUT):
    shutil.rmtree(OUT)
shutil.copytree(SRC, OUT)
wf_dir = os.path.join(OUT, 'Workflows')
for f in sorted(os.listdir(wf_dir)):
    fname, refs = process_flow(os.path.join(wf_dir, f))
    print(VARIANT, fname, sorted(refs))

sol = os.path.join(OUT, 'solution.xml')
x = open(sol, encoding='utf-8-sig').read()
x = x.replace('<Version>1.0.0.37</Version>', '<Version>%s</Version>' % NEW_VERSION)
open(sol, 'w', encoding='utf-8', newline='').write(x)

if SMTP:
    cz = os.path.join(OUT, 'customizations.xml')
    x = open(cz, encoding='utf-8-sig').read()
    add = ''
    for logical, disp, api in [(APPROVALS_REF, 'Aprobaciones AWGAutoconsolidaciones-8b3e4', 'shared_approvals'),
                               (SMTP_REF, 'SMTP AWGAutoconsolidaciones-5f1c2', 'shared_smtp')]:
        add += ('    <connectionreference connectionreferencelogicalname="%s">\n'
                '      <connectionreferencedisplayname>%s</connectionreferencedisplayname>\n'
                '      <connectorid>/providers/Microsoft.PowerApps/apis/%s</connectorid>\n'
                '      <iscustomizable>1</iscustomizable>\n'
                '      <promptingbehavior>0</promptingbehavior>\n'
                '      <statecode>0</statecode>\n'
                '      <statuscode>1</statuscode>\n'
                '    </connectionreference>\n') % (logical, disp, api)
    assert x.count('</connectionreferences>') == 1
    x = x.replace('  </connectionreferences>', add + '  </connectionreferences>')
    open(cz, 'w', encoding='utf-8', newline='').write(x)
