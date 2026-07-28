# AWG Ship With — el contacto se vuelve a pedir cada semana para el mismo Ship-to

## Síntoma

En la *Regional Team – Contacts Data Base* aparecen filas que son el **producto
cartesiano de Sold To × Ship To**:

| Sold to | Ship to Name | … | Email to |
|---|---|---|---|
| AWG - OKLAHOMA CITY | AWG - OKLAHOMA CITY | | mike.bourdelais@awginc.com |
| AWG - GREAT LAKES DIV | AWG - OKLAHOMA CITY | | dave.scanlan@awginc.com |
| AWG - SPRINGFIELD | AWG - OKLAHOMA CITY | | mike.bourdelais@awginc.com |
| AWG - OKLAHOMA CITY | AWG - GREAT LAKES DIV | | mike.bourdelais@awginc.com |
| … (9 combinaciones para 3 ship-to) | | | |

y aun así el flow sigue pidiendo el contacto **del mismo Ship-to todas las
semanas**.

## Cómo está armada la cadena hoy

```
[AWG - Ship with POs Repot]  (trigger: correo con adjunto .xlsx)
  Create_temp_report_file      -> ShipWith_Temp.xlsx
  List_contacts                -> Contacts Data Base (tabla 328EDC8D-…)
  Select_contacts              -> [{ name: <Ship to Name>, email: <Email to> }]
  Run_script_build_drafts      -> Office Script `shipWithReport`
                                  in : contactsJson, signatureHtml
                                  out: [{ to, cc, subject, htmlBody, matched }]
  For_each_draft
    Send_message
    If matched == false  -> Create_pending_item
                            Title = subject − "SHIP WITH NEEDED - "

[AWG ShipWith - Add Contact & Send]  (recurrence: 1 h)
  Get_pending_items   Status eq 'Pending' and ContactEmail ne null
  Add_all_contacts    -> AddRowV2  { "Ship to Name": Title, "Email to": ContactEmail }
  … agrupa por email, manda el correo, marca Status = Sent
```

## Dónde está el problema

Dos hallazgos, verificados sobre los `definition.json` exportados:

### 1. El Sold To **no** viene de la Contacts Data Base

`Select_contacts` (flow del reporte) mapea únicamente:

```json
{ "name": "@item()?['Ship to Name']", "email": "@item()?['Email to']" }
```

La columna *Sold to* nunca se le pasa al script. Y `Add_contact_row` (flow de
contactos) escribe únicamente `item/Ship to Name` y `item/Email to` — tampoco
toca la columna *Sold to*.

**Conclusión: el Sold To entra por el lado del reporte, dentro del Office Script
`shipWithReport`.** El script está construyendo su clave de agrupación / de
match combinando Sold To + Ship To en lugar de usar solo el Ship To. Eso explica
las dos mitades del síntoma:

- **Las combinaciones**: cada par (Sold To, Ship To) se trata como un
  destinatario distinto, así que 3 sold-to × 3 ship-to generan 9 grupos.
- **La repetición semanal**: cuando un Ship-to que ya tiene contacto aparece bajo
  un Sold To distinto, la clave combinada no coincide con ningún `name` de
  `contactsJson`, el script devuelve `matched = false`, y el flow vuelve a crear
  un item Pending pidiendo el contacto otra vez.

El `Title` que se guarda como *Ship to Name* sale de
`replace(subject, 'SHIP WITH NEEDED - ', '')`, es decir del `subject` que arma el
script. Si ese subject lleva la clave combinada, lo que queda escrito en la
Contacts DB tampoco es un Ship-to limpio, y nunca vuelve a hacer match.

### 2. El append a la Contacts DB no está protegido contra duplicados

`Add_all_contacts` llama a `AddRowV2` **incondicionalmente** para cada item
Pending. No consulta la tabla antes de escribir. Aunque se corrija la clave del
script, cualquier re-proceso (un reporte reenviado, un Pending resuelto dos
veces) sigue agregando filas repetidas.

## Correcciones

### A. Office Script `shipWithReport` — la causa raíz

El script debe:

1. Agrupar los POs y resolver el contacto **solo por Ship To**. El Sold To puede
   seguir mostrándose en el cuerpo del correo, pero no debe formar parte de la
   clave.
2. Normalizar ambos lados de la comparación antes de hacer match — `trim()` +
   `toUpperCase()` — para que "AWG - Great Lakes Div " y
   "AWG - GREAT LAKES DIV" cuenten como el mismo Ship-to.
3. Poner en `subject` el Ship-to limpio, porque ese texto es el que termina
   guardado como *Ship to Name* en la Contacts DB.

En pseudocódigo, el cambio es pasar de:

```ts
const key = `${soldTo} - ${shipTo}`;              // ❌ genera el cartesiano
const hit = contacts.find(c => c.name === key);
```

a:

```ts
const norm = (s: string) => (s ?? "").trim().toUpperCase();
const key  = norm(shipTo);                        // ✅ solo Ship-to
const hit  = contacts.find(c => norm(c.name) === key);
```

Este cambio **no está aplicado** en este repo: el script no viene en los
paquetes exportados (`Microsoft.Flow/flows/*/definition.json` solo referencia el
`scriptId` `01DYYVZGSTVFZKV6GUKFBLVFEWLC75I3O2`). Hay que exportarlo desde
Excel → Automate → `shipWithReport` para poder corregirlo con precisión.

### B. Flow `AWG ShipWith - Add Contact & Send` — aplicado

`flows/awg-shipwith-add-contact-send/definition.patched.json` agrega:

- **`List_existing_contacts`** — lee la Contacts Data Base (mismo file/table que
  usa el flow del reporte).
- **`Select_existing_ship_tos`** — proyecta las claves ya existentes
  normalizadas: `@toUpper(trim(coalesce(item()?['Ship to Name'], '')))`.
- **`If_ship_to_not_in_contacts`** — envuelve `Add_contact_row`; solo escribe si
  el Ship-to no está ya en la tabla.
- `item/Ship to Name` ahora se guarda con `@trim(...)`, y la columna *Sold to*
  se deja deliberadamente vacía para que el contacto quede identificado por el
  Ship-to únicamente.

El resto del flow (agrupación por email, envío, `Status = Sent`) queda igual.

### C. Limpieza de la Contacts Data Base

Las filas ya generadas siguen envenenando el match aunque se corrija el script.
Hay que dejar **una fila por Ship-to**:

1. Quedarse con un solo registro por *Ship to Name* (el correo correcto para ese
   destino).
2. Vaciar la columna *Sold to* en las filas que sobreviven.
3. Borrar las filas duplicadas del cartesiano.

En el ejemplo del screenshot, las 9 filas AWG colapsan a 3:

| Sold to | Ship to Name | Email to |
|---|---|---|
| *(vacío)* | AWG - OKLAHOMA CITY | mike.bourdelais@awginc.com |
| *(vacío)* | AWG - GREAT LAKES DIV | dave.scanlan@awginc.com |
| *(vacío)* | AWG - SPRINGFIELD | mike.bourdelais@awginc.com |

> Ojo con `AWG - GREAT LAKES DIV`: aparece dos veces con destinatarios distintos
> (`dave.scanlan@awginc.com` y `dave.scanlan@awginc.com; nam.le…`). Al colapsar
> hay que decidir cuál de los dos es el bueno.

## Cómo importar el flow corregido

1. Power Automate → **My flows** → *AWG ShipWith - Add Contact & Send* → **Export
   → Package (.zip)**, para tener respaldo.
2. Reemplazar `Microsoft.Flow/flows/<id>/definition.json` dentro del .zip por
   `flows/awg-shipwith-add-contact-send/definition.patched.json`.
3. **Import** el paquete como *Update* sobre el flow existente.
4. En el diseñador, volver a seleccionar Location / Document Library / File /
   Table en **List existing contacts** — los `drive`/`file`/`table` importados
   son IDs y el diseñador pide re-bindearlos.
5. Correr el flow una vez con un Pending de prueba y confirmar que no se agrega
   fila si el Ship-to ya existe.
