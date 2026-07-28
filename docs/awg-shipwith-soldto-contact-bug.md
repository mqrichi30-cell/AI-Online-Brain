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

## La Contacts Data Base NO se debe deduplicar

> Corrección: una versión anterior de este documento decía que había que
> colapsar la tabla a una fila por Ship-to. **Eso es incorrecto y destruiría
> datos.** Lo que sigue reemplaza esa instrucción.

El volcado real de la tabla (256 filas, ejecución del 2026-07-28 11:42) muestra
que **no hay duplicados**: la tabla está legítimamente llave-ada por
**Ship-to × Category**, y el CC cambia por categoría porque es el contacto
interno de P&G de esa categoría.

Columnas reales (16): `ION`, `OMA`, `Sales Office`, `Sales Group`,
`Sold to _x0023_`, `Ship to _x0023_`, `Sold to Name`, `Ship to Name`,
`Category`, `Buyer Name`, `Email to`, `CC`, `Other`, `Notes`, más
`@odata.etag` e `ItemInternalId` del conector.

```
AWG Great Lakes | FamilyCare  | dave.scanlan@awginc.com | CC: …;bright.al@pg.com
AWG Great Lakes | HairCare    | dave.scanlan@awginc.com | CC: …;grant.t.2@pg.com
AWG Great Lakes | OralCare    | dave.scanlan@awginc.com | CC: …;prettejohn.jl@pg.com
```

28 Ship-tos × ~9 categorías ≈ 256 filas. Todas las filas de un mismo Ship-to
comparten el mismo `Email to`, así que deduplicar por Ship-to para resolver el
destinatario (lo que hace el script) es correcto; **borrar filas de la tabla no
lo es** — se perdería el ruteo de CC por categoría, y la tabla la consumen otros
procesos además de este.

### Lo que sí hay que revisar

**1. El flow escribe filas incompletas.** `Add contact row` solo llena
`Ship to Name` y `Email to`; las otras 12 columnas quedan vacías, sin `Category`
ni `CC`. Eso ensucia una tabla maestra compartida. Decidir si el flow debe
seguir escribiendo ahí o registrar el contacto en otro lado.

**2. Los nombres de Ship-to no coinciden entre el reporte y la tabla.** En la
tabla: `AWG Great Lakes`, `AWG Gulf Coast`, `AWG Hernando`, `AWG - KANSAS CITY`,
`AWG Nashville`. En el reporte semanal: `AWG - GREAT LAKES DIV`,
`AWG - NASHVILLE`, `AWG - OKLAHOMA CITY`, `AWG - SPRINGFIELD`. `normName()` solo
normaliza mayúsculas y espacios — no reconcilia `AWG GREAT LAKES` con
`AWG - GREAT LAKES DIV`. **Esta es la causa candidata más fuerte de que el
contacto se pida cada semana**, y se resuelve alineando los nombres, no tocando
el código.

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

**Orden sugerido**: script → flow del reporte → flow de contactos → alinear los
nombres de Ship-to entre el reporte y la tabla → correr una semana en
observación.

## Tests

```bash
office-scripts/tests/run.sh                 # versión actual
git show <rev>:office-scripts/shipWithReport.ts > /tmp/old.ts
office-scripts/tests/run.sh /tmp/old.ts     # cualquier revisión anterior
```

Requiere `node` y `tsc`. El runner concatena el script con el harness (Office
Scripts comparte un solo scope de archivo), compila y ejecuta, y además verifica
que el script no dependa de `console`/DOM — cosas que Office Scripts no ofrece.
