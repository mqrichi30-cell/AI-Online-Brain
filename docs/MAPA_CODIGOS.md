# Mapa de codigos del diagrama — Automatizacion AWG Consolidaciones

Cada codigo `[F<nn>-A<seq>]` identifica un paso del diagrama y las acciones
reales del flujo de Power Automate que representa.

Carpeta de trabajo: Bandeja de entrada > "Marin, Cristhofer - AWG + Wakefern"
del buzon compartido pgcustservw2.im@pg.com
IA: ossgenai.im@pg.com — Seguimiento: lista SharePoint "Consolidation Tracking"
(sitio NACSO-RegionalVMI)

---

## FASE 00 · Router / Clasificador
Flujo: `GENERAL - 00 Email Router / Process Classifier`

| Codigo | Paso |
|---|---|
| F00-T | Entra un correo NUEVO a la carpeta del cliente |
| F00-A01 | Decision: es correo nuevo de cliente (no IA, no sistema, sin marcas AWG, sin "No Match") |
| F00-A02 | Decision: el asunto es "AWG - Ship With POs Report" |
| F00-A03–A05 | Decision: es correo de IA / sistema / Wakefern |
| F00-A06 | Decision: llego en los ultimos 60 minutos |
| F00-A07–A09 | Marca el correo LEIDO y le pone categorias "RPA" y "Cristhofer Marin" |
| F00-A10–A14 | Decision: esta conversacion ya fue ruteada antes (registro en SharePoint) |
| F00-A15–A16 | Crea registro de seguimiento con FlowId nuevo — Fase `00-RouterSent` |
| F00-A17–A19 | Responde el original a la IA pidiendo clasificar |
| F00-A20–A23 | Mueve el original a "GENERAL - Email Route - 00" |

Salidas sin continuidad: ignorado (no dispara), bloqueado (Ship With POs Report),
omitido por IA/sistema/ya ruteado, omitido por correo viejo, omitido duplicado.

---

## FASE 01 · Intake — lee la clasificacion de la IA
Flujo: `AWG CON - 01 Intake Authorized Sender v2 (FlowId)`

| Codigo | Paso |
|---|---|
| F01-T | Llega un correo a la carpeta del cliente |
| F01-A01–A02 | Decision: el remitente es la IA (ossgenai.im@pg.com) |
| F01-A03–A05 | Lee del cuerpo el FlowId oculto (AWG-FLOW-ID) y el resultado de clasificacion |
| F01-A06–A26, F01-A68–A71 | Decision: hay registro con ese FlowId en Fase `00-RouterSent` |
| F01-A27–A33 | Actualiza el registro con la clasificacion — Fase `00-RouterClassified` |
| F01-A34 | Decision: la clasificacion es "Consolidation" |
| F01-A37–A43 | Marca el original "RPA" + "Follow up" (no leido) y lo mueve a "AWG Complete > AWG - CON - 01" |
| F01-A35–A36 | Responde a la IA pidiendo la tabla Truck # / PO # — Fase `01-IntakeSent` |
| F01-A44–A49 | NO consolidacion: devuelve el original a la raiz como no leido y sin categorias. Fin |
| F01-A50–A67 | Sin registro (fallback): busca el original en varias carpetas, lo restaura y archiva la respuesta de la IA |

---

## FASE 02 · Recibe la tabla Truck/PO y pide Ride-With
Flujo: `AWG - CON - 02 Send Email to OSS Gen AI v3`

| Codigo | Paso |
|---|---|
| F02-T | Llega un correo de la IA a la carpeta del cliente |
| F02-A01–A04 | Decision: hay registro con ese FlowId en Fase `01-IntakeSent` |
| F02-A05–A10 | Decision: trae una tabla Truck/PO valida |
| F02-A11 | Decision: primera vez que se procesa esta consolidacion |
| F02-A12–A20 | Decision: se extrajeron 2 o mas numeros de PO |
| F02-A21 | Envia el correo "Ride-With" a pgcustservw2.im@pg.com |
| F02-A22–A28 | Guarda PO/SO en SharePoint — Fase `02-ConsolTableRcvd` |
| F02-A29–A41 | Menos de 2 PO: restaura el original, archiva y BORRA el registro. Fin |
| F02-A42–A54 | No era consolidacion: restaura el original, archiva y BORRA el registro. Fin |

---

## FASE 03 · Ride-With — arma el MOC y pide alinear la RDD
Flujo: `AWG CON - 03 Ride With Automation v2 (Guard)`

| Codigo | Paso |
|---|---|
| F03-T | Llega un correo al buzon compartido (se revisa cada 1 minuto) |
| F03-A01–A05, F03-A09–A12 | Decision: es la notificacion "Ride-With" |
| F03-A06–A08 | Decision: el FlowId aun NO fue procesado en esta fase |
| F03-A13–A18 | Extrae el PO principal y los PO que viajan con el |
| F03-A19 | Avanza el registro a Fase `03-RDDRequestSent` |
| F03-A20–A22 | Llena la plantilla MOC.xlsx con los PO (Office Script) |
| F03-A23–A26 | Envia "Awg Consolidation Request" con MOC.xlsx adjunto a nacsoshared.im@pg.com |
| F03-A27–A28 | Responde en el hilo a la IA pidiendo alinear los PO a la ultima RDD (POM CODE L3) |
| F03-A29–A31 | Mueve el Ride-With a "AWG Complete", lo marca leido y espera 1 hora |

---

## FASE 04 · Respuesta de RDD y seguimiento
Flujo: `AWG CON - 04 OSS Gen Response Tracking v2 (Guard)`

| Codigo | Paso |
|---|---|
| F04-T | Llega un correo al buzon compartido (cada 1 minuto) |
| F04-A01–A02 | Decision: trae la tabla de datos de la IA ("Pom code", "L3", "Current RDD") |
| F04-A03–A13 | Extrae la nueva RDD alineada y calcula fecha de trabajo = RDD − 10 dias |
| F04-A14–A16 | Decision: hay registro de esta conversacion en Fase `03-RDDRequestSent` |
| F04-A17 | Decision: hoy esta dentro de la ventana de trabajo (hoy >= RDD − 10 dias) |
| F04-A18–A19 | Status "Ready for OSS Follow Up" — Fase `05-FullTableRequestSent`, pide la tabla completa |
| F04-A20–A23 | Faltan mas de 10 dias: Status "Pending 10 Days Out" con fecha programada |
| F04-A24–A25 | Etiqueta el correo y lo mueve a "AWG Complete" |

---

## FASE 05 · Notificacion final de consolidacion
Flujo: `AWG CON - 05 Final Consolidation Notification v2 (Guard)`

| Codigo | Paso |
|---|---|
| F05-T | Llega un correo a la carpeta del cliente |
| F05-A01–A03 | Decision: hay registro de esta conversacion en Fase `05-FullTableRequestSent` |
| F05-A04–A08 | Decision: trae la tabla completa de la IA ("Formatted Table") |
| F05-A09–A21 | Extrae la tabla HTML y, del encabezado, el remitente original y destinatarios |
| F05-A22–A30 | Decision: algun "Collective Number" esta en blanco o difiere entre filas |
| F05-A31–A33 | Envia al remitente original "The following POs have been consolidated as requested" |
| F05-A34 | Avanza el registro a Fase `06-Consolidated` |
| F05-A35 | Todas las filas comparten Collective Number: avisa que debe trabajarse MANUALMENTE |
| F05-A36–A37 | Etiqueta la respuesta de la IA y la mueve a "AWG Complete" |

---

## FASE 06 · Barrido de flujos atascados
Flujo: `AWG CON - 06 Stuck Flow Sweep`

| Codigo | Paso |
|---|---|
| F06-T | Temporizador: cada 30 minutos |
| F06-A01 | Calcula la hora de corte: ahora − 5 horas |
| F06-A02 | Busca registros atascados (fase distinta de `06-Consolidated` y `99-ManualReview`) |
| F06-A03 | Decision: hay registros atascados |
| F06-A04 | Envia alerta "[AWG CON] Flujo de consolidacion atascado (+5h)" a pgcustservw2.im@pg.com |
| F06-A05 | Marca cada registro atascado como Fase `99-ManualReview` |
| F06-A06–A08 | Borra las filas ya terminadas (Fase `06-Consolidated`) |
