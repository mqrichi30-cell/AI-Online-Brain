# AWGAutoconsolidaciones 1.0.0.38

Dos copias de la solución `AWGAutoconsolidaciones_1_0_0_37_managed.zip`:

| Archivo | Envío de correos | Aprobaciones antes de enviar a OSSGenAI |
|---|---|---|
| `AWGAutoconsolidaciones_1_0_0_38_managed_Outlook.zip` | Igual que antes (Office 365 Outlook) | Outlook "Enviar correo con opciones" (Aprobar / Rechazar) |
| `AWGAutoconsolidaciones_1_0_0_38_managed_SMTP_Produccion.zip` | **SMTP** (`Send Email (V3)`) | Conector **Aprobaciones** ("Iniciar y esperar una aprobación", Approve / Reject) |

Cada copia viene en versión `managed` y `unmanaged`. El contenido es el mismo; solo cambia `<Managed>` en `solution.xml`. Hay que importar la versión que coincida con cómo está instalada la solución en el entorno. Si ya está como no administrada (por ejemplo en *Personal Productivity (default)*), usa `unmanaged`. Si no, aparece el error *"The solution is already installed on this system as an unmanaged solution…"*.

## Cambios en los 6 flujos (las dos copias)

- Se usa la estructura de `TryandCatchExample`:
  - Todos los `Initialize variable` (firma `SignatureHtml`, `varTrucksArray`, `FirstCollective`, etc.) van arriba, por fuera del Try y en cadena.
  - Todo el resto del flujo va dentro de `Scope_-_Try`, sin cambiar sus acciones ni su orden. Si una acción dependía de una variable, ahora depende de lo que esa variable esperaba antes.
  - `Scope_-_Catch` está copiado tal cual de la plantilla, sin cambios (`runAfter` del Try: Failed / TimedOut).
- Aprobación antes de cada correo a `ossgenai.im@pg.com`. Las 5 aprobaciones que ya tenían los flujos 00, 01, 03, 04 (tabla) y 04 (10 días) se mantienen y muestran el contenido que se va a enviar.
  - **Nuevo en el flujo 05:** si la fila de tracking no tiene `OriginalMessageId`, la respuesta usa el correo actual, que viene de OSSGenAI. Antes de cada una de las 2 respuestas, el flujo revisa a quién va a llegar (from / replyTo / CC). Si incluye OSSGenAI, pide aprobación. Si se rechaza o no hay respuesta en 24 h, no se envía nada, el correo vuelve a no leído y el flujo termina como Cancelled.
  - Al rechazar, el flujo termina como `Cancelled` y no como error, así que el Catch no se dispara.
- La versión pasa a 1.0.0.38.

## Cambios solo en la copia SMTP

- Ninguna acción de Outlook envía correo (lo comprueba `tools/validate.py`). Las acciones de Outlook que solo leen, mueven o marcan correos se mantienen.
- `Send an email (V2)` pasa a SMTP `Send Email (V3)`.
- `POST /reply` de Graph se reemplaza por 3 pasos:
  1. Un GET del mensaje original.
  2. Se arman los destinatarios con la misma regla de Graph: el `toRecipients` del request, o si no hay, `replyTo` / `from`. El asunto queda como `RE: …` y el cuerpo lleva el historial citado (From / Sent / To / Subject y el cuerpo original).
  3. El envío por SMTP.
- El borrador con `MOC.xlsx` más `/send` del flujo 03 pasa a ser un solo envío SMTP con el adjunto.
- Remitente: la variable `SmtpFrom` va arriba del Try, y al inicio del Try se llena con el buzón de la conexión de Outlook (`GET /me`). Así las respuestas de OSSGenAI siguen llegando a las carpetas que vigilan los triggers. En el flujo 04 se envía desde `pgcustservw2.im@pg.com`, igual que antes.
- Referencias de conexión nuevas: `awg_sharedsmtp_5f1c2` (SMTP) y `awg_sharedapprovals_8b3e4` (Aprobaciones). Para el texto de la aprobación se reutiliza la de Conversión de contenido, que ya existía.

## Revisar al importar en producción

1. **SMTP:** la cuenta SMTP necesita permiso para enviar como el buzón de `SmtpFrom` (y como `pgcustservw2.im@pg.com` en el flujo 04).
2. **Hilos:** el conector SMTP no permite poner `In-Reply-To` ni `References`, así que el correo sale como `RE: <asunto>` con el historial, pero Outlook puede no agruparlo en la misma conversación. Los flujos se enlazan principalmente por el FlowId que va en el cuerpo. Aun así, el flujo 01 busca el original por `conversationId` cuando no encuentra el tracking del router, así que hay que probar ese camino.
3. Después de importar, abrir una acción SMTP con "Peek code" y confirmar los nombres de los parámetros (`emailMessage/*`, `Attachments[].ContentData`) en el tenant.
4. El Catch de la plantilla arma el HTML del error, pero no lo envía (igual que en el ejemplo).

## Regenerar

```bash
python3 -I tools/build.py <solucion_extraida> <TryandCatch/definition.json> <salida> outlook|smtp
python3 -I tools/validate.py <solucion_extraida> <salida> outlook|smtp <TryandCatch/definition.json>
```
