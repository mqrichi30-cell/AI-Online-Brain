# AWG – Cancellation (flujo 00 + nuevo flujo RPA)

Objetivo:
1. **Modificar el flujo 00** (`GENERAL - 00 Email Router / Process Classifier`) para
   que OSSGenAI pueda identificar cuándo un correo es una **Cancellation**.
2. **Crear un flujo nuevo** que lea la respuesta, identifique las órdenes a cancelar
   y las envíe a un **RPA de cancelaciones** (Subject + contacto exactos).

---

## Parte 1 — Cambio en el flujo 00 (LISTO)

El 00 responde el hilo a `ossgenai.im@pg.com` con un prompt clasificador
(acción **`GENERAL - 00 HTTP Reply Original to OSSGenAI SameThread`**, campo `Body`,
dentro del `comment`). Hoy permite 3 valores de `ProcessType`:
`Consolidation`, `Wakefern Appointments`, `No Match`.

**Cambios aplicados** (ver `classifier-00-NEW-comment.html` y
`GENERAL-00-MODIFIED-definition.json`):

1. Se añadió una **regla de clasificación de Cancellation**:
   > **Cancellation rule:** Classify as Cancellation only when the latest actual
   > customer email contains BOTH: 1) an explicit customer request to cancel, void,
   > or delete one or more orders (cancel, please cancel, cancellation, void, delete
   > order, remove the order); and 2) at least one explicit PO or order number tied
   > to PO/order context (PO, PO#, PO #, purchase order, order, SO, sales order)…
   > A cancellation request is never a Consolidation. If either requirement is
   > missing, classify as No Match.

2. Se amplió la lista de valores de salida (de **tres** a **cuatro**):
   `Consolidation, Cancellation, Wakefern Appointments, No Match`.

Todo lo demás del prompt (READ SCOPE, exclusiones, formato de tabla de 1 columna /
1 fila, el `AWG-FLOW-ID`/`AWG-STEP` oculto) se mantiene **idéntico**.

### Cómo aplicarlo (solución **managed**)

> ⚠️ La solución exportada es **managed** (`Managed=1`, v1.0.0.1, publisher
> `awgnacso`). Una solución managed **no se edita en su entorno destino**. Aplica el
> cambio en el **entorno de origen / versión unmanaged** y reexporta, o:

- **Opción recomendada:** en tu entorno DEV, abre el flujo 00 (versión unmanaged),
  entra a la acción `GENERAL - 00 HTTP Reply Original to OSSGenAI SameThread` y
  reemplaza el texto del prompt por el de `classifier-00-NEW-comment.html`
  (es solo el `comment`; conserva el resto del JSON del `Body`). Sube la versión y
  reexporta la solución (managed) para PROD.
- **Alternativa:** importa `GENERAL-00-MODIFIED-definition.json` como definición del
  flujo (mismo flujo) si trabajas por definición.

> Nota de formato: dentro del flujo, los saltos del prompt van como `\n` **literal**
> (backslash-n), no como salto de línea real. El archivo `.json` ya lo trae así.

---

Además del cambio del prompt del 00, se **reestructuró** el prompt con un
**orden de decisión explícito** (exclusiones → Wakefern Appointments →
Cancellation → Consolidation → No Match), manteniendo el resto idéntico.

## Parte 2 — Handler de Cancellation (LISTO)

Paquete importable: **`AWG-Cancellation-Handler.zip`** (2 flujos), Office Script
`office-script/FillCancellationForm.ts`, y el template en `template/`.

```
[00] clasifica -> ProcessType = Cancellation
        │
        ▼
[CAN-01 Request]  (trigger: correo de ossgenai.im@pg.com, carpeta AWG)
        │   - aisla la respuesta de OSS (split '<div id=divrplyfwdmsg>')
        │   - si es clasificacion Cancellation, responde el hilo a OSSGenAI con
        │     un prompt de EXTRACCION (WORKFLOW=na_om_data_agent) que exige:
        │        ###CANCEL-START###
        │        SalesOrderNumber|ReasonCode|ReasonForCancellation
        │        ...
        │        ###CANCEL-END###
        │   - estampa AWG-STEP: CANCEL_EXTRACT (oculto)
        │
        │   (OSSGenAI tarda ~2-10 min)
        ▼
[CAN-02 Fill & Send]  (trigger: respuesta de OSSGenAI con el bloque)
        │   - normaliza HTML -> texto, extrae el bloque, parsea las filas
        │   - copia el template, corre FillCancellationForm (llena A/B/C desde fila 3)
        │   - adjunta el Excel y lo envia al RPA
        ▼
   Correo al RPA  ->  To: nacsoshared.im@pg.com
                       Subject: Order Cancellation Request
                       Adjunto: Order_Cancellation_Form.xlsx (lleno)
```

### El template (Order Cancellation Form)

| Col | Header exacto | Valores permitidos |
|-----|---------------|--------------------|
| A | `Sales Order Number ` | (los SO a cancelar) |
| B | `Reason Code (Choose from drop Down)` | `03`, `05`, `HD` |
| C | `Reason for cancelation (Choose from drop down) ` | `No longer needed`, `Duplicate order`, `Customer error`, `P&G error` |

Datos desde la fila 3. `FillCancellationForm` escribe solo VALORES, preservando
el layout y los dropdowns.

### Instalación (handler)

1. **Sube el template** `template/Order_Cancellation_Form.xlsx` a **OneDrive** de
   `marin.c@pg.com` en la raíz, con ese nombre exacto (path por defecto en el
   flujo: `/Order_Cancellation_Form.xlsx`; si va en otra carpeta, ajusta el
   parámetro `path` de `CAN_B_Get_Template_Content`).
2. **Crea el Office Script** `FillCancellationForm` en Excel Online (pega
   `office-script/FillCancellationForm.ts`) y copia su Script Id.
3. **Importa** `AWG-Cancellation-Handler.zip` (Import Package / Legacy). Mapea:
   Office 365 = `pgcustservw2.im@pg.com` (buzón AWG), OneDrive y Excel Online =
   `marin.c@pg.com`.
   > El `scriptId` de `CAN_B_Run_Fill_Script` apunta al script del 11s (referencia
   > válida en `marin.c`) para que el paquete **importe**. Tras importar,
   > **repunta** esa acción al `FillCancellationForm` real.
4. **Activa** los dos flujos.

### Anti-bucle / enrutamiento

- CAN-01 solo actúa si la parte superior de la respuesta de OSS contiene
  `processtype` + `cancellation` y **no** el bloque de extracción.
- CAN-02 solo actúa si el correo contiene `###CANCEL-START###`/`###CANCEL-END###`.
- El correo al RPA usa otro asunto (`Order Cancellation Request`) y no reentra por
  estos triggers (que escuchan a `ossgenai.im@pg.com`).

### Supuestos a VERIFICAR

- **Mapeo Reason Code ↔ Reason:** los dropdowns B (`03/05/HD`) y C
  (`No longer needed/...`) son independientes; **no conozco la correspondencia de
  negocio**. El prompt pide a OSSGenAI elegir ambos según el correo y, si no hay
  motivo, usa por defecto `03` + `No longer needed`. **Confírmame la tabla
  código↔motivo** y ajusto el prompt/default.
- **Identificador:** el RPA usa **Sales Order Number** (columna A). El prompt pide
  "sales order / order / PO number". Si necesitas conversión PO→SO, hay que añadir
  un paso.
- **Adjunto:** `CAN_B_Send_To_RPA` usa `Send an email (V2)` con
  `ContentBytes = @body('CAN_B_Get_Temp_Content')`. Si tu tenant exige base64
  explícito, se cambia a `base64(...)`.

El handler (CAN-01/02) es **independiente** (paquete legacy aparte) y no depende de
tocar la solución managed; el único cambio dentro de la solución es el prompt del
00 (Parte 1).
