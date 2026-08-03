# Instrucciones para que Claude en Chrome haga todo el trabajo

Este documento es para **Claude para Chrome** (la extensión que controla el
navegador). Con esto Claude ejecuta la automatización de punta a punta sin
Power Automate, y opcionalmente repara el flow para que vuelva a correr solo.

Hay **tres modos**. Empiece por el A.

| Modo | Para qué | Cuánto tarda |
|---|---|---|
| **A** | Ejecutar hoy la corrida que no salió | ~10 min |
| **B** | Reparar el flow para que vuelva a correr solo | ~20 min |
| **C** | Corrida diaria recurrente | 5 min/día |

---

## Antes de empezar

1. Instale **Claude para Chrome** desde `claude.ai/chrome`.
2. Inicie sesión en Chrome con la cuenta **pgcustservw2.im@pg.com** (o la que
   tenga acceso al buzón). Claude usa la sesión del navegador; no le dé
   contraseñas por chat.
3. Deje abiertas estas pestañas:
   - `https://outlook.office.com/mail/`
   - `https://make.powerautomate.com/`
4. Tenga a mano el archivo `office-script/ProcessWakefern11s_AUTOPUSH.ts` de
   este repo — Claude lo va a pegar en Excel.

> **Regla de seguridad que no se negocia:** el Modo A envía correo a
> **destinatarios externos de Wakefern**. Las instrucciones obligan a Claude a
> detenerse y mostrarle ambos borradores **antes** de enviar. No quite ese paso.

---

## MODO A — Ejecutar la corrida de hoy

Copie **todo** el bloque siguiente y péguelo en Claude para Chrome.

````text
Eres mi asistente de operaciones en Order Management. Vas a ejecutar
manualmente, dentro del navegador, la automatización "Wakefern - 11s block -
Autopush" que no corrió hoy. Trabaja paso a paso y confirma cada paso antes de
avanzar al siguiente.

REGLA CRÍTICA: no envíes ningún correo hasta que yo escriba la palabra
"ENVIAR". Los destinatarios son clientes externos reales.

=== PASO 1 — Localizar el reporte ===
1. Ve a https://outlook.office.com/mail/
2. Busca el correo más reciente cuyo asunto empiece con
   "Wholesale 2 Alerts Report" y venga de doom.im@pg.com.
3. Dime el asunto completo y la hora de recepción. Confirma que trae un adjunto
   llamado "Wholesale 2.xlsx".

=== PASO 2 — Abrir el adjunto en Excel para la web ===
1. En ese correo, en el adjunto "Wholesale 2.xlsx", usa la flecha del adjunto y
   elige "Guardar en OneDrive" (o "Cargar en OneDrive").
2. Cuando termine, abre el archivo en Excel para la web.
3. Confirma que el libro tiene una hoja llamada "Alerts".

=== PASO 3 — Cargar el Office Script ===
1. En Excel para la web, ve a la pestaña "Automatizar" (Automate).
2. Haz clic en "Nuevo script".
3. Borra TODO el contenido del editor de código.
4. Pega el código que te voy a dar en el siguiente mensaje.
5. Nombra el script "ProcessWakefern11s_AUTOPUSH" y guárdalo.
6. Avísame cuando esté guardado y te confirmo para ejecutarlo.

=== PASO 4 — Ejecutar y leer el resultado ===
1. Ejecuta el script con el botón "Ejecutar".
2. Abre el panel de resultado / registro y cópiame el objeto JSON completo que
   devolvió. Debe tener exactamente estos 7 campos:
   rowCount, ossHtmlRows, clientHtmlRows, clientSentenceDate,
   todayIso, tomorrowIso, pushGiIso
3. Dime cuántas órdenes encontró (rowCount) y a qué fecha se empuja el GI
   (pushGiIso).
4. Si rowCount es 0, detente por completo y avísame: hoy no hay nada que enviar.

=== PASO 5 — Redactar el correo 1 (OSS) — NO ENVIAR ===
Vuelve a Outlook y crea un mensaje nuevo, pero DÉJALO COMO BORRADOR:

  Para:    ossgenai.im@pg.com
  Asunto:  Wakefern RDD Change Request APPOINTMENT LATE - {todayIso}

  Cuerpo (formato HTML, respeta la tabla):

  Hello,

  Please apply this change:

  [tabla con dos columnas: "SO #" y "Push GI to", con las filas que vengan
   en el campo ossHtmlRows del JSON]

  Regards,
  pgcustservw2.im@pg.com
  NA Order Management | Regional

Muéstrame una captura o el texto del borrador y ESPERA mi confirmación.

=== PASO 6 — Redactar el correo 2 (Cliente Wakefern) — NO ENVIAR ===
Crea otro borrador:

  Para:    steve.salotti@wakefern.com; Mark.Kielczynski@wakefern.com;
           grocery_special_po_group@wakefern.com; anne.mucchiello@wakefern.com
  CC:      jackson.vs@pg.com
  Asunto:  APPOINTMENT NEEDED WAKEFERN

  Cuerpo (formato HTML):

  Hello team

  These POs had been assigned with an appointment too late, so we need to push
  to the soonest RDD possible that is {clientSentenceDate}, could you please
  provide an appointment?

  [tabla con dos columnas: "PO #" y "New RDD", con las filas que vengan en el
   campo clientHtmlRows del JSON]

  Regards,
  pgcustservw2.im@pg.com
  NA Order Management | Regional

Muéstrame el borrador y ESPERA.

=== PASO 7 — Envío (solo si yo escribo ENVIAR) ===
Cuando yo escriba exactamente "ENVIAR":
1. Envía primero el correo al OSS.
2. Confirma que salió.
3. Envía después el correo al cliente.
4. Confirma que salió.

=== PASO 8 — Cerrar el ciclo en el buzón ===
1. Marca como leído el correo original "Wholesale 2 Alerts Report...".
2. Muévelo a la carpeta "Misc Folder".
3. Borra de OneDrive el "Wholesale 2.xlsx" que subimos en el Paso 2.
4. Dame un resumen final: cuántas órdenes, a qué fecha se empujó el GI, y a
   quiénes se les envió.
````

**Segundo mensaje a Claude:** pegue el contenido completo de
`office-script/ProcessWakefern11s_AUTOPUSH.ts`.

---

## MODO B — Reparar el flow en Power Automate

Este modo deja la automatización corriendo sola otra vez. Pegue este bloque:

````text
Vas a reparar un flow de Power Automate que dejó de funcionar. Ve paso a paso y
confírmame cada hallazgo antes de cambiar nada.

=== PASO 1 — Diagnóstico (solo leer, no modificar) ===
1. Ve a https://make.powerautomate.com/
2. Abre "Mis flujos" y busca
   "Wakefern - 11s block - Autopush V14 WHOLESALE 2 ALERTS".
3. Dime si el flow está ENCENDIDO o APAGADO/SUSPENDIDO.
4. Abre el historial de ejecuciones de los últimos 28 días.
5. Busca la corrida del 03-08-2026 alrededor de las 07:03.
6. Reporta exactamente:
   - ¿Existe esa corrida? Si no existe, dímelo y detente ahí.
   - Si existe: ¿en qué estado terminó?
   - Si falló: ¿en cuál acción exactamente, y cuál es el mensaje de error
     completo?
No cambies nada todavía.

=== PASO 2 — Verificar el Office Script ===
La acción "Run script ProcessWakefern11s" apunta a un script guardado en
OneDrive con el id 01DYYVZGXRHHDY4HR32VGID4O5BXO2MOKB. Ese puntero se rompe si
el script se movió, renombró o borró.

1. Abre la acción "Run script ProcessWakefern11s" en el editor del flow.
2. Mira el campo "Script".
3. Dime si muestra un nombre de script válido o si muestra un error / id crudo /
   lista vacía.
4. Si está roto: abre el desplegable y dime qué scripts aparecen disponibles.

=== PASO 3 — Corregir el envío de correos ===
Este es el bug principal. Las acciones Create_OSS_Draft_HTTP,
Send_OSS_Draft_HTTP, Create_Customer_Draft_HTTP y Send_Customer_Draft_HTTP usan
"Send an HTTP request" del conector Office 365 Outlook con un URI ABSOLUTO
(https://graph.microsoft.com/v1.0/me/messages). Ese conector espera un URI
RELATIVO y le antepone su propia base, así que la llamada falla.

Hay dos formas de arreglarlo. Dime cuál prefieres antes de tocar nada:

  OPCIÓN 1 (mínima): en las 4 acciones, cambiar el URI absoluto por relativo:
     https://graph.microsoft.com/v1.0/me/messages   ->   me/messages
     y en las de envío:
     https://graph.microsoft.com/v1.0/me/messages/{id}/send  ->  me/messages/{id}/send

  OPCIÓN 2 (recomendada): borrar las 4 acciones HTTP y reemplazarlas por 2
     acciones nativas "Enviar un correo electrónico (V2)" de Office 365 Outlook.
     Elimina el Graph y el JSON armado a mano. Yo te doy los valores exactos de
     asunto, destinatarios y cuerpo cuando me digas que vas por aquí.

=== PASO 4 — Probar sin mandarle nada al cliente ===
1. Antes de cualquier prueba, cambia el destinatario del correo al cliente por
   mi propia dirección. Dime cuando esté cambiado.
2. Usa "Probar" > "Manualmente" o reenvía el correo del reporte al buzón.
3. Reporta el resultado acción por acción.
4. Solo cuando la corrida salga verde y los correos de prueba lleguen bien,
   me avisas para restaurar los destinatarios reales de Wakefern.

Nunca restaures los destinatarios reales por tu cuenta: eso lo autorizo yo.
````

---

## MODO C — Corrida diaria

Una vez que el Modo A funcionó al menos una vez, para el día a día basta con:

````text
Ejecuta la corrida diaria de Wakefern 11s siguiendo el procedimiento que ya
conoces:

1. Busca en Outlook el "Wholesale 2 Alerts Report" de HOY.
2. Sube el adjunto a OneDrive y ábrelo en Excel para la web.
3. Pestaña Automatizar > ejecuta el script "ProcessWakefern11s_AUTOPUSH".
4. Muéstrame el JSON del resultado.
5. Prepara los DOS borradores (OSS y Cliente) con esos datos.
6. NO ENVÍES nada. Espera a que yo escriba ENVIAR.
7. Después de enviar: marca leído, mueve a "Misc Folder", borra el temporal de
   OneDrive.
````

---

## Si Claude se traba

| Síntoma | Qué pedirle a Claude |
|---|---|
| No encuentra la pestaña "Automatizar" en Excel | *"Estoy en Excel para la web, no en la app de escritorio. Si no ves 'Automatizar', dime qué pestañas ves."* |
| El script devuelve `rowCount: 0` | *"Muéstrame las filas de la hoja Alerts donde 'Sold to Name' contenga WAKEFERN y 'DB' sea 11, con su columna 'Good Issue'."* |
| El script da error de compilación | *"Pégame el mensaje de error exacto y la línea."* Suele ser que quedó código viejo sin borrar en el editor. |
| Las tablas del correo salen sin bordes | *"Pega el cuerpo como HTML, no como texto plano. Usa el modo de edición HTML del compositor."* |
| Se salta el paso de confirmación | Corte y repita: *"Detente. No envíes nada. Muéstrame el borrador primero."* |

---

## Verificación independiente (opcional, sin navegador)

Si quiere comprobar que Claude en Chrome hizo bien las cuentas, corra localmente
el mismo cálculo sobre el `.msg` original:

```bash
pip install olefile openpyxl
python3 tools/run_local.py "Wholesale_2_Alerts_Report__08032026_Time_07_03.msg" --today 2026-08-03
```

Genera `output/script_result.json` (lo que debe devolver el script),
`output/emails.json` y `output/preview.html`. Si el JSON de Claude coincide con
`script_result.json`, la corrida es correcta.
