# AWG Autoconsolidaciones (Power Automate)

- `AWGAutoconsolidaciones/` – solución desempaquetada (fuente).
- `AWGAutoconsolidaciones_1_0_0_35.zip` – solución lista para importar.
- `add_ossgenai_approvals.py` – script que agregó la aprobación (se corre sobre la 1.0.0.34).

## Aprobación antes de enviar a OSSGenAI (1.0.0.35)

Todo correo a `ossgenai.im@pg.com` pasa ahora por un correo de aprobación
("Send email with options") con dos botones: **Aprobar** / **Rechazar**.
El correo muestra el contenido exacto que se enviará.

| Flujo | Envío a OSSGenAI protegido | Dónde está la aprobación |
|---|---|---|
| GENERAL - 00 | `GENERAL_-_00_HTTP_Reply_Original_to_OSSGenAI_SameThread` | justo antes del envío |
| CON - 01 | `AWG_-_CON_-_01_HTTP_Reply_Current_OSS_to_OSSGenAI_SameConversation` | justo antes del envío |
| CON - 03 | `AWG_-_CON_-_03_HTTP_Send_RDD_Reply_By_Immutable_ID` | antes de actualizar tracking / enviar MOC |
| CON - 04 | `AWG_-_CON_-_04_Send_HTTP_OSS_Response` y `..._1` | antes de actualizar el item de SharePoint |

CON - 02 y CON - 05 no envían correos a OSSGenAI, no cambiaron.

- **Aprobar**: el flujo sigue igual que antes.
- **Rechazar** (o sin respuesta en 24 h, o si la aprobación falla): no se envía nada,
  no se notifica a nadie, el correo original vuelve como **no leído** a la carpeta
  `Marín, Cristhofer - AWG + Wakefern` y la ejecución termina como *Cancelled*.

La solicitud se envía al buzón de la conexión de Outlook del flujo (`GET /me`).
Para mandarla a otra dirección, cambia el campo **To** de la acción `*_Aprobacion_Enviar_Solicitud`.
