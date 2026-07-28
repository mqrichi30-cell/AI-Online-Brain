# AWG Ship With — el contacto se vuelve a pedir cada semana para el mismo Ship-to

## Síntoma

En la *Regional Team – Contacts Data Base* se acumulan filas que parecen
combinaciones de Sold To × Ship To, y aun así el flow sigue pidiendo el contacto
**del mismo Ship-to todas las semanas**.

## Cómo está armada la cadena

```
[AWG - Ship with POs Repot]  (trigger: correo con adjunto .xlsx)
  Create_temp_report_file      -> ShipWith_Temp.xlsx
  List_contacts                -> Contacts Data Base (tabla 328EDC8D-…)
  Select_contacts              -> [{ name: <Ship to Name>, email: <Email to> }]
  Run_script_build_drafts      -> Office Script `shipWithReport`
  For_each_draft
    Send_message
    If matched == false  -> Create_pending_item

[AWG ShipWith - Add Contact & Send]  (recurrence: 1 h)
  Get_pending_items   Status eq 'Pending' and ContactEmail ne null
  Add_all_contacts    -> AddRowV2  { "Ship to Name": Title, "Email to": ContactEmail }
  … agrupa por email, manda el correo, marca Status = Sent
```

## Causa raíz

El Sold To **no** viaja desde la Contacts Data Base: `Select_contacts` mapea solo
`Ship to Name` y `Email to`, y `Add_contact_row` escribe solo esas dos columnas.
El problema estaba en el Office Script, y son **dos defectos independientes**,
ambos reproducidos con `office-scripts/tests/run.sh` contra el script original:

### 1. Una fila duplicada con el email vacío borra un contacto bueno

```ts
for (const c of contacts) {
    if (c && c.name) {
        contactMap[normName(c.name)] = String(c.email || "").trim();   // ❌
    }
}
```

La asignación es incondicional y **la última fila gana**. Como la tabla tiene
varias filas por Ship-to, basta con que una de ellas tenga el nombre lleno y el
email vacío para que `contactMap["AWG - GREAT LAKES DIV"]` quede en `""` →
`matched = false` → el flow crea otro Pending → **te vuelve a pedir el contacto
de un Ship-to que ya lo tenía**.

Contra el script original (test T1):

```
T1  duplicate contact row with a BLANK email must not erase a good contact
  FAIL matched              expected true   actual false
  FAIL to                   expected "dave.scanlan@awginc.com"
                            actual   "pgcustservw2.im@pg.com"
  FAIL pending rows created expected []     actual ["AWG - GREAT LAKES DIV"]
```

Y con filas duplicadas que sí tienen email, "la última gana" también elige un
destinatario **arbitrario**: en el screenshot `AWG - GREAT LAKES DIV` aparece con
`dave.scanlan@` y con `mike.bourdelais@`, y el que quede de último es el que se
usa.

### 2. Un Ship-to escrito de dos formas produce un Title combinado

```ts
const key = matched ? email.toLowerCase() : "__unmatched__" + normName(b.shipTo);
…
if (groups[key].names.indexOf(b.shipTo) < 0) groups[key].names.push(b.shipTo);  // ❌ raw
…
subject: "SHIP WITH NEEDED - " + g.names.join("/"),
```

La clave del grupo está **normalizada** pero en `names` se guarda el texto
**crudo**. Si el reporte trae `AWG - Great Lakes Div` y `AWG - GREAT LAKES DIV`,
las dos caen en el mismo grupo y el subject queda:

```
SHIP WITH NEEDED - AWG - Great Lakes Div/AWG - GREAT LAKES DIV
```

El flow del reporte hacía `item/Title = replace(subject, 'SHIP WITH NEEDED - ', '')`,
así que ese string **combinado** se guardaba tal cual como un único
`Ship to Name`. Esa fila no puede coincidir nunca con un Ship-to real → el
Ship-to queda pidiéndose para siempre, y la tabla acumula filas que parecen
combinaciones.

Contra el script original (test T2):

```
T2  two spellings of ONE unmatched Ship-to must yield ONE clean Pending row
  FAIL pending rows created
       expected ["AWG - Great Lakes Div"]
       actual   ["AWG - Great Lakes Div/AWG - GREAT LAKES DIV"]
```

Lo mismo aplica a los grupos con `matched = true`, que se agrupan **por email**:
varios Ship-tos con el mismo destinatario comparten un subject `A/B`. Ese caso no
llegaba a la Contacts DB (no genera Pending), pero confirmaba que el subject no
sirve como identificador.

## Correcciones aplicadas

### A. `office-scripts/shipWithReport.ts`

1. **`contactMap` no destructivo** — se ignoran las filas con email vacío y gana
   el primer email no vacío, así una fila duplicada o a medio llenar ya no puede
   borrar un contacto existente.
2. **Identidad = Ship-to normalizado** — los bloques se pliegan primero en una
   entrada por Ship-to (`normName`), con un nombre canónico único. Los Ship-tos
   *sin* contacto **nunca** se agrupan entre sí: uno por draft, garantizando un
   solo Pending limpio por Ship-to. Los que *sí* tienen contacto se siguen
   agrupando por email, para que cada persona reciba un solo correo.
3. **Campo `shipTo` en el payload** — el draft ahora expone el Ship-to canónico
   aparte del `subject`, para que el flow no tenga que recortar texto.
4. **El Sold To no puede leerse como Ship-to** — `col()` resuelve primero la
   columna Sold-to y la marca como bloqueada, de modo que el fallback por
   coincidencia parcial de `cShipTo` no pueda caer en ella. Si no hay columna
   Ship-to, el script devuelve `[]` en vez de agrupar todo bajo una clave falsa.

> Nota de alcance: (1) y (2) son las causas confirmadas del síntoma. (4) es
> endurecimiento — con los headers de ejemplo el script original también
> resolvía bien la columna (test T3 pasa en ambas versiones); queda para que un
> cambio de headers en el reporte no reintroduzca el problema.

Comportamiento preservado: el HTML del correo, el subject, el CC, la
consolidación por destinatario y el ruteo de los no encontrados al buzón
compartido no cambian (tests T4 y T5).

### B. `flows/awg-shipwith-pos-report/definition.patched.json`

`Create_pending_item` ahora toma el nombre del campo dedicado:

```diff
- "item/Title": "@replace(items('For_each_draft')?['subject'], 'SHIP WITH NEEDED - ', '')"
+ "item/Title": "@items('For_each_draft')?['shipTo']"
```

### C. `flows/awg-shipwith-add-contact-send/definition.patched.json`

`Add_all_contacts` llamaba a `AddRowV2` sin consultar la tabla, así que cualquier
reproceso duplicaba filas. Se agregó:

- **`List_existing_contacts`** — lee la Contacts Data Base (mismo file/table que
  usa el flow del reporte).
- **`Select_existing_ship_tos`** — claves existentes normalizadas:
  `@toUpper(trim(coalesce(item()?['Ship to Name'], '')))`.
- **`If_ship_to_not_in_contacts`** — envuelve `Add_contact_row`; solo escribe si
  el Ship-to no está ya en la tabla.
- `item/Ship to Name` se guarda con `@trim(...)`, y la columna *Sold to* se deja
  vacía a propósito.

## Limpieza de la Contacts Data Base

Las filas ya generadas siguen envenenando el match aunque el código esté
corregido: mientras exista una fila con nombre y email vacío, o un Title
combinado, el Ship-to correspondiente se seguirá pidiendo. Hay que dejar **una
fila por Ship-to**:

1. Borrar las filas cuyo `Ship to Name` contenga `/` (los Titles combinados).
2. Borrar las filas con `Ship to Name` lleno y `Email to` vacío.
3. Quedarse con un solo registro por `Ship to Name` y vaciar la columna
   *Sold to* en el que sobrevive.

En el ejemplo del screenshot, las 9 filas AWG colapsan a 3:

| Sold to | Ship to Name | Email to |
|---|---|---|
| *(vacío)* | AWG - OKLAHOMA CITY | mike.bourdelais@awginc.com |
| *(vacío)* | AWG - GREAT LAKES DIV | dave.scanlan@awginc.com |
| *(vacío)* | AWG - SPRINGFIELD | mike.bourdelais@awginc.com |

> `AWG - GREAT LAKES DIV` aparece con dos destinatarios distintos
> (`dave.scanlan@awginc.com` y `dave.scanlan@awginc.com; nam.le…`). Al colapsar
> hay que decidir cuál es el bueno — el código ya no lo elige al azar, pero
> tampoco puede adivinar.

## Cómo desplegar

**Office Script** — Excel → Automate → `shipWithReport` → pegar
`office-scripts/shipWithReport.ts` → Save. El script devuelve un campo nuevo
(`shipTo`), así que **hay que desplegarlo junto con el flow del reporte**: si se
actualiza solo el flow, `shipTo` llega vacío y los Pending se crean sin nombre.

**Flows** — para cada uno: Export → Package (.zip) como respaldo, reemplazar
`Microsoft.Flow/flows/<id>/definition.json` por el `definition.patched.json`
correspondiente, e Import como *Update* sobre el flow existente. En el diseñador
hay que volver a seleccionar Location / Document Library / File / Table en
**List existing contacts** — los `drive`/`file`/`table` importados son IDs y el
diseñador pide re-bindearlos.

**Orden sugerido**: script → flow del reporte → flow de contactos → limpieza de
la tabla → correr una semana en observación.

## Tests

```bash
office-scripts/tests/run.sh                 # versión actual
git show <rev>:office-scripts/shipWithReport.ts > /tmp/old.ts
office-scripts/tests/run.sh /tmp/old.ts     # cualquier revisión anterior
```

Requiere `node` y `tsc`. El runner concatena el script con el harness (Office
Scripts comparte un solo scope de archivo), compila y ejecuta, y además verifica
que el script no dependa de `console`/DOM — cosas que Office Scripts no ofrece.
