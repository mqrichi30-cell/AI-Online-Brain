# Especificación — Automatización de correos AWG (Consolidaciones)

Documento de referencia para analizar y modificar los flujos de Power Automate del proceso
de consolidaciones AWG. Generado a partir de los `definition.json` exportados de Power Automate
(paquetes legacy ZIP). Refleja el estado actual de los flujos al 2026-07-02.

## 1. Contexto general

- **Buzón compartido:** `pgcustservw2.im@pg.com`
- **Carpeta de trabajo:** Bandeja de entrada > `Marín, Cristhofer - AWG + Wakefern`
- **Subcarpetas usadas:**
  - `GENERAL - Email Route - 00` (donde Fase 00 estaciona el correo original mientras la IA clasifica)
  - `AWG Complete` (archivo de correos ya procesados)
  - `AWG Complete > AWG - CON - 01` (archivo específico de la fase de intake)
- **IA (agente de correo):** `ossgenai.im@pg.com` — se le escribe respondiendo en el mismo hilo con un
  prompt; contesta en el mismo hilo. El prompt siempre inicia con `FOR ROUTING USE WORKFLOW=na_om_data_agent (AskOM)`.
- **Buzón de consolidaciones (humanos/robot):** `nacsoshared.im@pg.com`
- **Correlación entre fases:** cada caso lleva un GUID `FlowId` que viaja escondido en el cuerpo de los
  correos en un párrafo invisible: `AWG-FLOW-ID: <guid> | AWG-STEP: <paso>` (fuente 1px blanca).
- **Lista de seguimiento (SharePoint):** sitio `https://pgone.sharepoint.com/sites/NACSO-RegionalVMI`,
  lista con id `70ed8abc-c524-4142-9d8a-86f70ced426c`. Columnas usadas: `Title`, `FlowId`,
  `OriginalMessageId`, `OriginalConversationId`, `OriginalSubject`, `Status`, `Phase`, `PhaseSince`,
  `POs`, `SOs`, `WorkDate`, `NewRDD`.
- **Plantilla Excel:** `MOC.xlsx` en SharePoint `NACSO-RegionalVMI /Shared Documents/General/Customer Documents/AWG-VMC/`,
  se llena/limpia con dos Office Scripts.

### Máquina de estados (columna `Phase` en SharePoint)

```
00-RouterSent → 00-RouterClassified → 01-IntakeSent → 02-ConsolTableRcvd
   → 03-RDDRequestSent → 05-FullTableRequestSent → 06-Consolidated (fila se borra luego)
99-ManualReview = atascado, marcado por el barrido (Fase 06 / flujo Sweep)
```

## 2. Flujos (uno por fase del diagrama)

### FASE 00 — `GENERAL - 00 Email Router / Process Classifier`
Archivo: `GENERAL00EmailRouterProcessClassifier_*.zip`

- **Disparador:** nuevo correo en la carpeta del cliente. Condiciones del trigger (todas):
  remitente NO contiene `ossgenai.im@pg.com` ni `nacsoshared.im@pg.com`; asunto NO contiene
  `awg-flow-id:` ni `awg-step:`; categorías NO contienen `no match`. Concurrencia = 1.
- **Lógica:**
  1. IF asunto contiene "awg - ship with pos repot/report" → **fin** (bloqueado).
  2. ELSE compone un texto con subject/from/to/cc/body y IF contiene alguno de: `wakefern`,
     `ossgenai.im@pg.com`, `final state results`, `for routing use workflow=na_om_data_agent`,
     `awg-flow-id:`, `awg-step:`, `email not found for`, `process has completed execution`,
     `no unread emails found`, `please do not respond directly`, `system-generated email`,
     `nacsoshared.im@pg.com` → **fin** (correo de IA/sistema/Wakefern, omitido).
  3. ELSE IF el correo NO llegó en los últimos 60 min → **fin** (viejo). *(Ojo: la expresión
     real es `ticks(received) < ticks(utcNow()-60min)` en la rama THEN vacía; la rama útil es el ELSE.)*
  4. Obtiene el **ImmutableId** del mensaje (Graph, `Prefer: IdType="ImmutableId"`).
  5. **PATCH al original:** `isRead=true`, categorías `["RPA","Cristhofer Marín"]`.
     ⚠️ CORRECCIÓN PENDIENTE #1: este paso debe moverse para DESPUÉS de responder a la IA (paso 8).
  6. Chequeo de duplicado en SharePoint: busca fila con mismo `OriginalMessageId` O
     `OriginalConversationId` O `OriginalSubject`. Si existe → **fin** (ya ruteado).
  7. Crea fila de seguimiento: `Title=GENERAL-00-<FlowId>`, `Status=Flow0-PromptSentToIA`,
     `Phase=00-RouterSent`, `FlowId=guid()` nuevo.
  8. **Responde el correo original en el mismo hilo** a `ossgenai.im@pg.com` con el prompt de
     clasificación (Consolidation / Wakefern Appointments / No Match) + marca oculta
     `AWG-FLOW-ID | AWG-STEP: PROCESS_CLASSIFIER`.
  9. Mueve el original a `GENERAL - Email Route - 00` (lookup dinámico de carpetas vía Graph).

### FASE 01 — `AWG CON - 01 Intake Authorized Sender v2 (FlowId)`
Archivo: `AWGCON01IntakeAuthorizedSenderv2FlowId_*.zip`

- **Disparador:** nuevo correo en la carpeta del cliente (sin filtro en trigger).
- **Lógica:**
  1. IF remitente == `ossgenai.im@pg.com`, si no → **fin**.
  2. Extrae del cuerpo el `FlowId` (parseando después de `AWG-FLOW-ID:`) y el `ProcessType`
     (busca en bodyPreview/body: "wakefern appointments" → Wakefern Appointments; "no match" → No Match;
     "consolidation" → Consolidation; default No Match).
  3. Resuelve IDs de carpetas (raíz cliente, Email Route 00, AWG Complete, AWG - CON - 01) y localiza
     el "OSS actual" (la respuesta de la IA) buscando por conversationId en 4 carpetas.
  4. Busca fila SharePoint `FlowId eq X and Phase eq '00-RouterSent'`.
     - **Sí existe:** actualiza fila (`Status=<ProcessType>`, `Phase=00-RouterClassified`); resuelve el
       correo original (por conversationId dentro de Email Route 00, excluyendo remitentes IA/robot).
       - IF `ProcessType == Consolidation`:
         responde en la misma conversación a la IA pidiendo la **tabla Truck # / PO #**
         (marca `AWG-STEP: CONSOLIDATION_CONFIRM`); avanza fila a `Phase=01-IntakeSent`;
         PATCH original: categorías `["RPA","Follow up"]`, `isRead=false`; mueve original a `AWG - CON - 01`;
         marca la respuesta IA `isRead=false` + `["RPA","Follow up"]` y la mueve a `AWG - CON - 01`.
       - ELSE (no consolidación): mueve el original de vuelta a la carpeta raíz del cliente,
         lo marca no leído y sin categorías; marca la respuesta IA leída sin categorías y la mueve a `AWG - CON - 01`. **Fin.**
     - **No existe (fallback):** busca el original por conversationId en Email Route 00 + raíz + AWG Complete
       (excluyendo remitentes IA/robot). Si lo encuentra: si está en AWG Complete lo deja; si está en la raíz
       lo marca no leído; si está en otra carpeta lo mueve a la raíz y lo marca no leído. La respuesta IA se
       marca leída y se mueve a `AWG - CON - 01`. **Fin** (no continúa la automatización).

### FASE 02 — `AWG - CON - 02 Send Email to OSS Gen AI v3 (No Consolidation Restore)`
Archivo: `AWGCON02SendEmailtoOSSGenAIv3NoConsolidationRestore_*.zip`

- **Disparador:** nuevo correo en la carpeta del cliente con `from=ossgenai.im@pg.com`.
- **Lógica:**
  1. Extrae `FlowId` del cuerpo; busca fila `FlowId eq X and Phase eq '01-IntakeSent'`.
     Si no hay → **Terminate(Cancelled)** (no está en el flujo).
  2. Normaliza el cuerpo, recorta solo el mensaje más reciente (`<div id=divrplyfwdmsg>`).
  3. IF el mensaje actual contiene `<table>` + `truck` + `po` y NO contiene `Customer name`,
     `Shipment Number`, `Pom code`, `SO #`, `Delivery Number`, `Is not a consolidation Regards`:
     - IF NO hay marcas de proceso previo en el hilo (`Confirm if this email is a consolidation request`,
       `Is not a consolidation`, `scheduled to ride with another load`, `AWG_CONSOLIDATION_REQUEST`):
       - Parsea filas de la tabla → arrays `varPOs`, `varSOs`.
       - IF `len(varPOs) >= 2`:
         crea **draft de respuesta al hilo del original** (createReply por ImmutableId del original guardado
         en SharePoint), le pone el HTML "ride with" (Main PO = primer PO; Riding-With = resto;
         marca `AWG-STEP: RIDE_WITH_REQUEST`), destinatario `pgcustservw2.im@pg.com`, y lo envía;
         actualiza fila (`POs`, `SOs`, `Phase=02-ConsolTableRcvd`); mueve la respuesta IA a `AWG Complete`.
       - ELSE (<2 POs): restaura el original a la carpeta raíz (no leído), marca la respuesta IA leída,
         la mueve a `AWG Complete`, **borra la fila** de SharePoint y termina.
  4. ELSE (no es tabla válida / dice "Is not a consolidation"): igual que el caso <2 POs:
     restaura original a raíz (no leído), archiva respuesta IA en `AWG Complete`, **borra la fila**, fin.

### FASE 03 — `AWG CON - 03 Ride With Automation v2 (Guard)`
Archivo: `AWGCON03RideWithAutomationv2Guard_*.zip`

- **Disparador:** nuevo correo en el buzón compartido `pgcustservw2.im@pg.com`,
  carpeta del cliente, polling cada 1 minuto.
- **Lógica:**
  1. Obtiene ImmutableId; extrae `FlowId` del cuerpo.
  2. Guard anti-duplicado: busca `FlowId eq X and Phase eq '02-ConsolTableRcvd'`; el flujo solo
     continúa si el filtro local da 0... *(nota: la condición está invertida respecto a lo intuitivo:
     entra a procesar cuando `length(filtered)==0` es FALSE-branch del nombre "Guard_Not_In_Flow";
     revisar el JSON si se modifica)*.
  3. IF el cuerpo contiene "please be advised that the following purchase orders will ride together..."
     o "this is an automatic notification that the following purchase order(s) are scheduled to ride with another load:":
     - IF NO contiene "final state results" ni "delivery number":
       - HTML→texto; extrae `Main PO Number (The load):` y `Riding With PO(s)...:` → lista All_POs.
       - PATCH fila SharePoint: `POs`, `Phase=03-RDDRequestSent`.
       - Office Script 1: escribe los PO en `MOC.xlsx`; espera 10 s; lee el archivo; lo adjunta en base64.
       - Crea y envía correo nuevo **"Awg Consolidation Request"** con `MOC.xlsx` a `nacsoshared.im@pg.com`.
       - Office Script 2: limpia `MOC.xlsx`.
       - Responde en el hilo a la IA pidiendo **alinear RDD** de todos los PO con POM CODE L3
         (`ProcessType: AWG_RDD_ALIGNMENT`, incluye `OriginalConversationId`).
       - Mueve el correo ride-with a `AWG Complete`, lo marca leído, espera 1 h (Delay) y termina.

### FASE 04 — `AWG CON - 04 OSSGen Response Tracking v2 (Guard)`
Archivo: `AWGCON04OSSGenResponseTrackingv2Guard_*.zip`

- **Disparador:** nuevo correo en buzón compartido, carpeta del cliente, cada 1 minuto.
- **Lógica:**
  1. IF el cuerpo contiene `Customer Item #`, `Pom code`, `L3` y `Current RDD` (tabla de datos de la IA):
  2. Extrae la nueva RDD del HTML (parsing frágil: `<td>21/`…) y la convierte a ISO;
     calcula `WorkDate = NewRDD − 10 días`.
  3. Busca fila `OriginalConversationId eq <convId> and Phase eq '03-RDDRequestSent'`;
     si 0 → **Terminate(Cancelled)**.
  4. IF `utcNow() >= NewRDD − 10 días` (ya dentro de la ventana):
     - PATCH fila: `Status=Ready for OSS Follow Up`, `WorkDate`, `NewRDD`, `Phase=05-FullTableRequestSent`.
     - Responde a la IA pidiendo la **tabla completa de consolidación** (columnas: Truck #, PO #, SO #,
       Customer name, RDD, Collective Number, Shipment Number, Delivery Number).
  5. ELSE (faltan >10 días):
     - PATCH fila: `Status=Pending 10 Days Out`, `WorkDate`, `NewRDD` (la `Phase` NO cambia).
     - Responde a la IA avisando la fecha programada de trabajo (`WorkDate`).
  6. En ambos casos: PATCH categorías del correo `["RPA","Follow up","Cristhofer Marín"]`
     y lo mueve a `AWG Complete`.

### FASE 05 — `AWG CON - 05 Final Consolidation Notification v2 (Guard)`
Archivo: `AWGCON05FinalConsolidationNotificationv2Guard_*.zip`

- **Disparador:** nuevo correo en la carpeta del cliente.
- **Lógica:**
  1. Busca fila `OriginalConversationId eq <convId> and Phase eq '05-FullTableRequestSent'`;
     si 0 → **Terminate(Cancelled)**.
  2. IF el cuerpo contiene `Formatted Table`, `PO #`, `SO #`, `Customer name`, `RDD`, `Collective Number`:
     - Extrae la tabla HTML y sus filas; extrae del encabezado reenviado el remitente original
       y los destinatarios del To (arma CC dinámico excluyendo `pgcustservw2.im@pg.com` y al remitente).
     - Recorre filas: toma el `Collective Number` (columna 6); marca `CollectiveBlankFound` si alguno
       está vacío y `CollectiveMismatch` si difieren entre filas.
     - IF `CollectiveBlankFound OR CollectiveMismatch`:
       envía correo (SendEmailV2) **al remitente original**, CC dinámico, asunto `RE: <subject>`,
       cuerpo "The following POs has been consolidated as requested" + tabla;
       PATCH fila `Phase=06-Consolidated`.
     - ELSE: envía correo a `pgcustservw2.im@pg.com` — "already within the 10-day window,
       needs to be reviewed and worked manually" + tabla.
     - Luego (ambas ramas): PATCH categorías del correo IA `["RPA","Follow Up","AWG OSSGenAI Processed"]`
       y lo mueve a `AWG Complete`.
  3. ELSE (no es la tabla completa) → fin sin acciones.

### FASE 06 — `AWG CON - 06 Stuck Flow Sweep`
Archivo: `AWGCON06StuckFlowSweep_*.zip`

- **Disparador:** Recurrencia cada 30 minutos (no depende de correos).
- **Lógica:**
  1. `cutoff = utcNow() − 5 h`.
  2. Busca filas con `Phase ne '06-Consolidated' and Phase ne '99-ManualReview' and PhaseSince lt cutoff` (top 100).
  3. Por cada atascada: envía alerta a `pgcustservw2.im@pg.com`
     (asunto `[AWG CON] Flujo de consolidacion atascado (+5h) - Fase <X>`, con FlowId, fase, PhaseSince,
     POs, asunto original, conversationId) y PATCH `Phase=99-ManualReview`.
  4. Busca filas `Phase eq '06-Consolidated'` (top 200) y **borra** cada una.

## 3. Correcciones pendientes (el usuario irá agregando más)

1. **FASE 00:** mover el paso "marcar como leído + categorías (`GENERAL_-_00_HTTP_Patch_Original_Read_Tags`)"
   para que se ejecute DESPUÉS de `GENERAL_-_00_HTTP_Reply_Original_to_OSSGenAI_SameThread`
   (hoy corre antes del chequeo de duplicado). Al reordenar, cuidar las dependencias `runAfter`:
   `Compose_Original_Conversation_ID` hoy corre después del PATCH; deberá colgarse de
   `Compose_Original_Immutable_ID` directamente.
2. *(pendiente — el usuario indicó que tiene más correcciones por dictar)*

## 4. Entregables esperados al aplicar correcciones

- Editar los `definition.json` dentro de los paquetes ZIP exportados (formato legacy de Power Automate:
  `manifest.json` en la raíz + `Microsoft.Flow/flows/<guid>/definition.json|apisMap.json|connectionsMap.json`).
- Reempaquetar cada flujo modificado como ZIP **manteniendo la estructura interna exacta** para que
  Power Automate lo acepte en *Importar → paquete legacy*. No cambiar GUIDs ni connection references.
- Actualizar el diagrama si el cambio afecta el orden visible: ejecutar `python3 generar_flujo_drawio.py`
  (script en la raíz del repo) que produce `Flujo_Correos_AWG.drawio`; editar las listas `pX_main`,
  `pX_sides`, `pX_edges` del script (formato: nodos con id/tipo/etiqueta/alto; las ramas laterales se
  alinean automáticamente a su rombo).

## 5. Convenciones del diagrama (draw.io)

- Óvalos = inicio (verde) / fin (rojo); rectángulos azules = pasos; rombos ámbar = decisiones;
  rectángulos naranjas = ramas alternas; óvalos grises = correo ignorado/descartado.
- Cada fase vive en un contenedor pastel; las fases se conectan con flechas moradas punteadas
  que van por los pasillos entre contenedores (waypoints fijos).
- Lenguaje no técnico, enfocado en "qué pasa con el correo".
