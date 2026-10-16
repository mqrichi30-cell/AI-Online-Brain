# AI-Online-Brain

Repositorio de trabajo para la solución de Power Automate **AWG Autoconsolidaciones** (publisher `awgnacso`, prefijo `awg`). Los entregables son zips de solución de Dataverse, listos para importar.

## Estructura

```
solutions/AWGAutoconsolidaciones/
  AWGAutoconsolidaciones_1_0_0_38_managed_Outlook.zip         # envíos por Office 365 Outlook
  AWGAutoconsolidaciones_1_0_0_38_unmanaged_Outlook.zip
  AWGAutoconsolidaciones_1_0_0_38_managed_SMTP_Produccion.zip # envíos por SMTP + conector Aprobaciones
  AWGAutoconsolidaciones_1_0_0_38_unmanaged_SMTP_Produccion.zip
  README.md                                                   # qué cambió y qué revisar al importar
  tools/build.py      # genera las variantes a partir del zip 1.0.0.37 + plantilla TryandCatch
  tools/validate.py   # valida runAfter, referencias, Catch intacto, aprobaciones, cero envíos Outlook en SMTP
```

Los zips `managed` y `unmanaged` solo se diferencian en `<Managed>` de `solution.xml`. Hay que importar el que coincida con el modo en que la solución ya está instalada en el entorno. Si no coincide, sale el error *"already installed … as an unmanaged solution"*.

## Flujos (6)

| Flujo | Trigger | Qué hace |
|---|---|---|
| GENERAL - Classifier - 00 - Send Classify Prompt | Correo nuevo en la carpeta "Marín, Cristhofer - AWG + Wakefern" | Espera horario laboral (L-V 5am-3pm CR) → filtros (reportes, sistema, >24h, ya enrutado) → clasifica por reglas (Consolidation / Cancellation / Return / ZQs / No Match) → registro SharePoint → **aprobación** → reply a OSSGenAI → mueve a "GENERAL - Email Route - 00" |
| AWG - Consolidation - 01 - Send Con Prompt | Respuesta de OSSGenAI | Si es consolidación: **aprobación** → reply a OSSGenAI pidiendo Truck # / PO # |
| AWG - Consolidation - 02 - Format AI Output | Respuesta de OSSGenAI | Separa por truck, crea tracking hijo, envía a pgcustservw2 |
| AWG - Consolidation - 03 - Send MOC Form | Buzón compartido | Run script Excel MOC → envía Consolidation Request (MOC.xlsx) a nacsoshared → **aprobación** → reply RDD a OSSGenAI |
| AWG - Consolidation - 04 - Query Con Status | Buzón compartido | Lee RDD → **aprobación** (tabla o aviso 10 días) → reply a OSSGenAI desde pgcustservw2 |
| AWG - Consolidation - 05 - Confirm Con Client | Respuesta de OSSGenAI con "Formatted Table" | Responde al cliente; **aprobación condicional** si la respuesta llegaría a OSSGenAI |

## Reglas de la solución (no romper)

1. **Try/Catch en todos los flujos:**
   - Todos los `Initialize variable` (firma `SignatureHtml`, etc.) van arriba, por fuera del Try y encadenados.
   - Todo lo demás va dentro de `Scope_-_Try`.
   - `Scope_-_Catch` es una copia exacta de la plantilla TryandCatchExample y **no se modifica**. Su `runAfter` es Try: Failed / TimedOut.
2. **Antes de cualquier correo a `ossgenai.im@pg.com` debe haber una aprobación**, y el correo de aprobación muestra el contenido que se va a enviar. Si se rechaza o vence (24 h), no se envía nada y el flujo termina con un `Terminate` Cancelled; el Catch no se dispara.
3. **Variante SMTP (producción):**
   - Ninguna acción de Outlook puede *enviar* correo. Leer, mover o marcar con Graph `HttpRequest` sí está permitido.
   - Los envíos usan SMTP `SendEmailV3`.
   - Las aprobaciones usan `shared_approvals` `StartAndWaitForAnApproval`.
   - El remitente sale de la variable `SmtpFrom`, que se llena con `GET /me`. En el flujo 04 el remitente es `pgcustservw2.im@pg.com`.
4. Para cambiar algo se edita `tools/build.py`, se regenera, se corre `tools/validate.py` y se reempaquetan los 4 zips. No se editan los JSON a mano.

## Cambios ya hechos y pendientes conocidos

- Hecho: el flujo 00 descartaba los correos del fin de semana. `Guard_Only_New_Email` comparaba contra `utcNow()` después de la espera de horario laboral. Ahora compara contra `trigger()?['startTime']`.
- Pendiente: **el usuario no recibe los correos de aprobación**. Hay que diagnosticarlo en el historial de corridas de Power Automate (ver `NEXT_SESSION.md`).
- SMTP no permite `In-Reply-To` / `References`, así que las respuestas pueden quedar fuera de la conversación en Outlook. En el flujo 01 conviene probar el camino que busca por `conversationId`.

## Convenciones

- Idioma de trabajo con el usuario: español.
- Commits y PRs en la rama `claude/try-catch-email-approval-7y2iaq`. PR: mqrichi30-cell/AI-Online-Brain#6.
