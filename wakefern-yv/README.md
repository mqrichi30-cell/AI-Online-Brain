# Wakefern – YV block automation

Automatiza el manejo de bloqueos **YV** (columna `Del Blk Indicator`) del Regional
Report, en dos etapas, siguiendo el mismo patrón del flujo existente
`Wakefern - 11s block - Autopush`.

El paquete importable (`Wakefern-YV-and-11s-Autopush.zip`) contiene **3 flujos**:

| Flujo | Estado | Qué hace |
|-------|--------|----------|
| `Wakefern - 11s block - Autopush V13 ...` | **sin cambios** | El flujo 11s original, intacto. |
| `Wakefern - YV block - Stage 1 Request Qty` | **nuevo** | Detecta líneas YV y pide cantidades a OSSGENAI. |
| `Wakefern - YV block - Stage 2 Update Qty` | **nuevo** | Lee la respuesta de OSSGENAI y pide la actualización. |

---

## Flujo del proceso YV

```
Regional Report llega
        │
        ▼
[Stage 1]  Office Script ProcessWakefernYV
        │   - filtra Del Blk Indicator == YV
        │   - filtra RDD entre HOY y HOY+10 (inclusive); fuera de rango = NO se trabaja
        ▼
   Correo a ossgenai.im@pg.com
   "Wakefern YV Qty Confirmation Request - <fecha>"
   Tabla SO # / Line # / RDD  +  exige respuesta en formato EXACTO
        │
        │   (OSSGENAI tarda ~2–10 min en responder)
        ▼
[Stage 2]  Se dispara con la respuesta de ossgenai.im@pg.com
        │   - extrae el bloque entre marcadores
        │   - lee SO#, Line#, Confirmed Qty
        ▼
   Correo a ossgenai.im@pg.com
   "Wakefern YV Quantity Update - <fecha>"
   Tabla SO # / Line # / New Target QTY (= Confirmed Qty)
```

---

## El formato EXACTO exigido a OSSGENAI (clave para que Stage 2 lea siempre igual)

Stage 1 le pide a OSSGENAI que responda incluyendo **este bloque literal**,
mismo encabezado, mismas 4 columnas, mismo orden, separado por `|`, una línea por
fila, **texto plano** (sin negritas, sin tablas, sin HTML) entre los marcadores:

```
###PGYV-START###
SO#|Line#|Target QTY|Confirmed Qty
<SO#>|<Line#>|<Target QTY>|<Confirmed Qty>
###PGYV-END###
```

Ejemplo de respuesta válida:

```
###PGYV-START###
SO#|Line#|Target QTY|Confirmed Qty
1234567|10|500|480
1234567|20|300|300
###PGYV-END###
```

Cómo lo parsea Stage 2 (robusto ante el HTML del correo):

1. Normaliza el HTML → texto (convierte `<br>`, `</div>`, `</p>` en saltos de línea).
2. Toma el texto entre el **primer** `###PGYV-START###` y el **primer** `###PGYV-END###`
   (así ignora la copia citada del correo original que suele venir debajo).
3. Divide por líneas, se queda con las que tienen `|` y descarta encabezado/marcadores.
4. Por cada línea: `split('|')` → posición 0 = SO#, 1 = Line#, **3 = Confirmed Qty**.
5. Arma la tabla de actualización y la envía a OSSGENAI.

> Como el segundo correo tiene otro asunto (`Wakefern YV Quantity Update`), las
> respuestas de OSSGENAI a ese correo **no** vuelven a disparar Stage 2 (que filtra
> por `Wakefern YV Qty Confirmation Request`). No hay bucle.

---

## Pasos de instalación

1. **Crear el Office Script** en Excel Online:
   - Abre un Regional Report en Excel Online → *Automate* → *New Script*.
   - Pega el contenido de `office-script/ProcessWakefernYV.ts`.
   - En el bloque `COLS` ajusta los nombres de columna **exactos** de tu reporte
     (`SO #`, `Line #`, `RDD`). El de bloque ya es `Del Blk Indicator`.
   - **Save**. Abre el script y copia su **Script Id** (aparece en la URL / al
     usarlo desde Power Automate).

2. **Importar el paquete** `Wakefern-YV-and-11s-Autopush.zip` en Power Automate
   (*My flows → Import → Import Package (Legacy)*). Durante la importación mapea
   las 3 conexiones (Office 365 Outlook = `pgcustservw2.im@pg.com`, OneDrive y
   Excel Online) a las mismas que usa el flujo 11s.

3. **Poner el Script Id**: en el flujo *Stage 1*, acción
   `Run_script_ProcessWakefernYV`, reemplaza el placeholder
   `PASTE_YV_OFFICE_SCRIPT_ID_HERE` por el id real del paso 1
   (formato `ms-officescript%3A%2F%2F...`, igual que el 11s).

4. **Revisar destinatarios/carpetas** (ya vienen configurados, ajústalos si hace falta):
   - Stage 1 trigger: correos de `pgcustservw1/w2.im@pg.com`, asunto `Regional Report`,
     misma carpeta que monitorea el 11s.
   - Stage 2 trigger: correos de `ossgenai.im@pg.com`, asunto contiene
     `Wakefern YV Qty Confirmation Request`, carpeta `Inbox`.

5. **Activar** los dos flujos nuevos.

---

## Supuestos tomados (fáciles de cambiar)

- **Arquitectura:** 2 flujos YV nuevos; el flujo 11s queda intacto. Stage 1 y
  Stage 2 están separados porque Stage 2 se dispara con la respuesta de OSSGENAI.
- **Identificador de línea:** `SO #` + `Line #`.
- **Actualización:** se le pide a OSSGENAI poner `Target QTY = Confirmed Qty`.
- **Sin correo al cliente** (a diferencia del 11s): el YV solo intercambia con OSSGENAI.
- **Ventana RDD:** hoy ≤ RDD ≤ hoy+10 (inclusive). Fila sin fecha o fuera de rango = no se trabaja.
- Stage 2 mapea **Confirmed Qty = columna 4** del bloque (índice 3 tras `split('|')`).

Si algún supuesto no aplica (p. ej. el identificador debe ser `PO #`, o la
actualización es otra regla), se ajustan puntualmente el Office Script y/o las
expresiones `concat`/`Select` del flujo.

## Depuración

Cada flujo trae acciones `Debug_*` (Compose) para inspeccionar en el historial de
ejecución: el JSON parseado del script, el JSON exacto de cada borrador, el bloque
extraído de la respuesta, y las salidas de error de las llamadas HTTP a Graph.
