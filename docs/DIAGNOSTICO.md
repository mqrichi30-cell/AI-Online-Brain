# Diagnóstico — por qué no corrió la automatización del 03-08-2026

Analizado a partir de los dos archivos entregados:

| Archivo | Qué es |
|---|---|
| `Wholesale_2_Alerts_Report__08032026_Time_07_03.msg` | El correo real que debía disparar el flow |
| `Wakefern11sblockAutopushV14WHOLESALE2ALERTS_20260803160020.zip` | La exportación del flow de Power Automate |

---

## 1. El correo estaba bien. El problema es el flow.

Verifiqué el `.msg` contra cada condición del trigger y de los filtros. **Todo pasa:**

| Condición del flow | Valor exigido | Valor real del correo | |
|---|---|---|---|
| Trigger `from` | `doom.im@pg.com;pgcustservw1.im@pg.com;pgcustservw2.im@pg.com` | `doom.im@pg.com` | ✅ |
| Trigger `subjectFilter` | `Wholesale 2 Alerts Report` | `Wholesale 2 Alerts Report - 08-03-2026 Time:07:03` | ✅ |
| Trigger `fetchOnlyWithAttachment` | `true` | 1 adjunto | ✅ |
| `Only_Wholesale_2_Alerts_Subject` | `startsWith(subject, 'Wholesale 2 Alerts Report')` | coincide | ✅ |
| `Filter_XLSX_Attachments` → `.xlsx` | termina en `.xlsx` | `Wholesale 2.xlsx` | ✅ |
| `Filter_XLSX_Attachments` → `isInline == false` | no inline | `AttachFlags=0`, `AttachmentHidden=0`, `RenderingPosition=-1` → **no inline** | ✅ |
| Office Script V14 lee la hoja `Alerts` | hoja `Alerts` | la única hoja del libro se llama `Alerts` | ✅ |
| `If_Rows_Found` → `rowCount > 0` | > 0 | **7 órdenes califican** | ✅ |

Las 7 órdenes que el flow debía procesar (WAKEFERN, DB `11`, Good Issue hoy o mañana):

| SO # | PO # | Good Issue |
|---|---|---|
| 2066923910 | 01737647 | 2026-08-03 |
| 2066923927 | 01737651 | 2026-08-03 |
| 2066923928 | 01737653 | 2026-08-03 |
| 2066923929 | 01737648 | 2026-08-03 |
| 2066923911 | 01737654 | 2026-08-04 |
| 2066923925 | 01737646 | 2026-08-04 |
| 2066923926 | 01737652 | 2026-08-04 |

**Conclusión: el insumo era correcto. La falla está dentro del flow.**

---

## 2. Defectos encontrados en el ZIP

### 🔴 Defecto A — El Office Script no existe dentro del ZIP

El paquete exportado contiene **exactamente 5 archivos**:

```
manifest.json
Microsoft.Flow/flows/manifest.json
Microsoft.Flow/flows/6d2947b7-.../definition.json
Microsoft.Flow/flows/6d2947b7-.../apisMap.json
Microsoft.Flow/flows/6d2947b7-.../connectionsMap.json
```

Ninguno es el script. El flow solo guarda un **puntero**:

```
scriptId = ms-officescript%3A%2F%2Fonedrive_business_itemlink%2F01DYYVZGXRHHDY4HR32VGID4O5BXO2MOKB
```

Ese id es un item de OneDrive de **marin.c@pg.com**. Los Office Scripts viven en
`/Documents/Office Scripts/` del usuario y **nunca se exportan junto al flow**.

Consecuencia: si el script se renombró, se movió, se borró, o el flow se importó
en otra cuenta, `Run_script_ProcessWakefern11s` falla y **toda la cadena se
detiene ahí** — no se envía ningún correo, no se mueve el mensaje.

Este es el defecto que hace que el ZIP, por sí solo, **no sea reproducible**.

→ **Corregido:** el script está reconstruido en
[`office-script/ProcessWakefern11s_AUTOPUSH.ts`](../office-script/ProcessWakefern11s_AUTOPUSH.ts).

---

### 🔴 Defecto B — URI absoluto en las 4 acciones HTTP de Outlook

El V14 crea y envía los correos con la acción *Send an HTTP request* del
conector Office 365 Outlook (`operationId: HttpRequest`). Ese conector espera
un **URI relativo** de Graph y le antepone su propia base. El V14 le pasa un
URI **absoluto**:

```
Uri = https://graph.microsoft.com/v1.0/me/messages
```

El resultado efectivo es una URL duplicada (`.../https://graph.microsoft.com/...`)
que devuelve 400/404.

Lo revelador: **la descripción de la propia acción dice lo contrario del valor
que tiene configurado.**

> `Create_OSS_Draft_HTTP` → *"V11 correction: Office 365 Outlook HttpRequest URI
> must be `me/messages` (no `/` and no `/v1.0`)."*
>
> …pero el valor real es `https://graph.microsoft.com/v1.0/me/messages`.

Es decir: en la V11 alguien ya había encontrado y corregido este bug, y en la
V13 se revirtió a URL absoluta ("V13 uses absolute Graph URLs"). Las 4 acciones
(`Create_OSS`, `Send_OSS`, `Create_Customer`, `Send_Customer`) tienen el mismo
problema.

→ **Corregido:** las 4 acciones HTTP se reemplazaron por **2 acciones nativas
`Send an email (V2)`**. No usan Graph, no arman JSON a mano, no manejan ids de
borrador. Menos piezas, menos superficie de falla.

---

### 🟠 Defecto C — JSON armado a mano con `concat()` + `json()`

El V14 construye el payload como una cadena de texto e inyecta ahí los datos
del script:

```
concat('{"subject":"...","body":{"contentType":"HTML","content":"...',
       body('Parse_Script_Result')?['ossHtmlRows'], '..."}}')
```

Cualquier comilla, backslash o salto de línea dentro de `ossHtmlRows`,
`clientHtmlRows` o `clientSentenceDate` rompe el `json()` y tumba el flow.
Es frágil por construcción.

→ **Corregido:** `Send an email (V2)` recibe el HTML como parámetro normal, sin
JSON intermedio. Además el script nuevo escapa el HTML (`escapeHtml`).

---

### 🟠 Defecto D — `Move_email` y `Mark_as_read` estaban dentro del `Foreach`

En el V14 ambas acciones viven dentro de `Apply_to_each_XLSX_Attachment`. Si el
correo llegara con **dos** adjuntos `.xlsx`, la segunda iteración intenta mover
un mensaje que ya se movió → falla.

→ **Corregido:** se movieron fuera del `Foreach`, al nivel del `If` de asunto.

---

### 🟡 Defecto E — El Excel temporal nunca se borra

`Create_Temp_Excel_File` escribe en la **raíz** de OneDrive
(`folderPath: "/"`) un archivo por cada ejecución, y nadie lo borra. Con una
corrida diaria eso son ~250 archivos al año en la raíz.

→ **Corregido:** ahora escribe en `/Wholesale2AlertsTemp` y se agregó
`Delete_Temp_Excel_File`, que corre pase lo que pase con el `If`.

---

### 🟡 Defecto F — Sin reintento en `Run_script_ProcessWakefern11s`

Ejecutar un Office Script contra un archivo **recién creado** en OneDrive falla
de forma intermitente porque el archivo todavía está bloqueado o sin indexar.

→ **Corregido:** se agregó `retryPolicy` exponencial (4 intentos, 10 s).

---

## 3. Cómo confirmar cuál de los defectos disparó la falla del 03-08

El historial de ejecuciones es la única fuente que lo dice con certeza. En
Power Automate:

1. `make.powerautomate.com` → **Mis flujos** → *Wakefern - 11s block - Autopush V14*
2. Pestaña **Historial de ejecuciones (28 días)**, buscar la corrida del
   **03-08-2026 ~07:03**
3. Interpretación:

| Lo que ve | Defecto responsable |
|---|---|
| **No hay ninguna corrida** | El trigger no disparó: flow apagado/suspendido, conexión de Outlook caída, o el correo no entró a `Inbox` (regla de Outlook lo movió antes) |
| Corrida en rojo, falla en `Run_script_ProcessWakefern11s` | **Defecto A** — el script no se encuentra |
| Corrida en rojo, falla en `Create_OSS_Draft_HTTP` / `Send_..._HTTP` | **Defecto B** — URI absoluto |
| Corrida en rojo, falla en `Parse_Script_Result` | El script devolvió algo distinto al contrato de 7 campos |
| Corrida en **verde** pero sin correos | `Filter_XLSX_Attachments` devolvió 0, o `rowCount = 0` |

Los defectos A y B son los únicos que explican una falla **total y silenciosa**
como la que ocurrió, y ambos están corregidos en la V15.

---

## 4. Regla de negocio que hay que confirmar

El script V14 original no está disponible, así que la regla exacta de cuánto se
empuja el GI **no se puede recuperar del ZIP**. La reconstruí con el criterio
más común y la dejé como constantes editables al inicio del `.ts`:

```ts
const PUSH_GI_BUSINESS_DAYS_AHEAD = 2;   // nuevo GI  = hoy + 2 días hábiles
const RDD_OFFSET_BUSINESS_DAYS    = 1;   // nuevo RDD = nuevo GI + 1 día hábil
```

Con el correo del 03-08-2026 (lunes) eso da:

- **Push GI a:** `2026-08-05` (miércoles)
- **Nuevo RDD:** `2026-08-06` (jueves) → *"Thursday, August 6, 2026"*

⚠️ **Si su regla real es otra, cambie esos dos números antes de enviar nada.**
Es el único punto del entregable que no pude derivar de los archivos.
