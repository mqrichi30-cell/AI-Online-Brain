# AI-Online-Brain

## scripts/Export-SharedMailboxEmails.ps1

Exporta a CSV los correos de los últimos N días desde carpetas concretas de un buzón
compartido de Microsoft 365, leyendo directamente de Exchange Online vía Microsoft Graph
(sin Outlook COM ni caché OST, por lo que los datos coinciden con New Outlook).

### Requisitos

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -Scopes "Mail.Read.Shared"
```

### Uso

```powershell
# Valores por defecto: buzón pgcustservw2.im@pg.com, 15 días,
# carpetas "AWG Complete" y "Wakefern Complete", CSV en Descargas
.\scripts\Export-SharedMailboxEmails.ps1

# Otra ventana de tiempo y otra ruta de salida
.\scripts\Export-SharedMailboxEmails.ps1 -Days 30 -OutputPath "C:\Temp\Ultimos30.csv"

# Si un nombre de carpeta no coincide, imprime el árbol completo del buzón
.\scripts\Export-SharedMailboxEmails.ps1 -ListFoldersOnly

# Incluir también las subcarpetas de cada carpeta destino
.\scripts\Export-SharedMailboxEmails.ps1 -IncludeSubfolders
```

### Por qué es más rápido que el enfoque anterior

| | Antes | Ahora |
|---|---|---|
| Filtro de fecha | En el cliente, tras descargar toda la carpeta | `$filter=receivedDateTime ge ...` en el servidor |
| Propiedades | Todo el mensaje | `$select` con 12 campos |
| Tamaño de página | 10 (valor por defecto) | `$top=999` |
| Acumulación | `$array += ...` (O(n²)) | `List[object]` (O(n)) |
| Throttling (429) | Aborta la ejecución | Reintento respetando `Retry-After` |
| CSV bloqueado en Excel | Se pierde todo el trabajo | Guarda en un archivo con marca de tiempo |

Con `-All` y sin filtro, Graph paginaba de 10 en 10 sobre el historial completo de la
carpeta: una carpeta con 40.000 correos son 4.000 peticiones para acabar quedándose con
los 15 días. Con el filtro y `$top=999` son solo las páginas del rango pedido.

### Columnas del CSV

`Received`, `ReceivedUtc`, `Subject`, `From`, `FromName`, `To`, `Cc`, `HasAttachments`,
`IsRead`, `Importance`, `Folder`, `ConversationId`, `InternetMessageId`, `WebLink`.

### Notas

- El recurso `message` de Graph v1.0 no expone una propiedad `size`, por eso no aparece
  en el `$select` (era la causa del error `Could not find a property named 'size'`).
- La coincidencia de nombres de carpeta ignora mayúsculas y espacios repetidos, para que
  `"Marín, Cristhofer - AWG + Wakefern"` no falle por un espacio de más.
- El CSV se escribe en UTF-8 con BOM para que Excel muestre bien los acentos.

## scripts/Export-And-PivotByHour.ps1

Hace los dos pasos anteriores de una sola vez: exporta los correos de los **últimos 6
meses** y arma el Excel con la tabla dinámica de correos por hora, sin pasar por
`Alt+F11`.

```powershell
.\scripts\Export-And-PivotByHour.ps1              # 6 meses
.\scripts\Export-And-PivotByHour.ps1 -Months 12
.\scripts\Export-And-PivotByHour.ps1 -SkipExcel   # solo el CSV
```

Genera `CorreosUltimos6Meses.csv` y `CorreosUltimos6Meses.xlsx` en Descargas, y abre el
Excel al terminar.

Diferencia importante con la macro VBA: **la hora se calcula en PowerShell** a partir del
`DateTime` real que devuelve Graph, no parseando texto en Excel. No depende del formato
regional ni de cómo Excel interprete la columna `Received`. El CSV ya sale con las
columnas `Hora`, `Rango horario`, `Fecha` y `DiaSemana` listas.

Antes de tocar Excel imprime el histograma por hora en la consola, así el resultado está
disponible aunque la automatización COM falle. Si Excel no está o la automatización se
rompe, avisa y deja el CSV igual en lugar de perder la corrida.

`Export-And-PivotByHour-OneLiner.ps1` es el mismo flujo condensado en un bloque plano,
pensado para pegar directamente en la consola de PowerShell sin guardar archivo.

## scripts/Get-HourlyBreakdown.ps1

Desglose de correos por hora **sin abrir Excel**. Alternativa a la tabla dinámica cuando
Excel bloquea el guardado.

```powershell
.\scripts\Get-HourlyBreakdown.ps1
.\scripts\Get-HourlyBreakdown.ps1 -PorDiaSemana
```

Imprime la tabla con histograma en consola y escribe un CSV **ya cruzado** (una fila por
franja horaria, una columna por carpeta, más `Total` y `Porcentaje`, con fila de totales).
Ese CSV es el resultado final: no hace falta armar ninguna dinámica encima. Con
`-PorDiaSemana` agrega el cruce hora × día de la semana.

Deriva la hora de la columna `Hora` si está, y si no la calcula desde `Received`. Las filas
sin hora reconocible se cuentan y se reportan, no se descartan en silencio.

## Etiquetas de confidencialidad y el guardado desde Excel

En entornos con etiquetado obligatorio (Microsoft Purview / Azure Information Protection),
Excel abre un diálogo modal al guardar un archivo nuevo. Desde automatización eso da dos
fallos, ambos malos:

- Con `DisplayAlerts = $false`, Excel contesta el diálogo solo, **cancela el guardado y no
  lanza error**: el script reporta éxito y no hay archivo.
- Con `DisplayAlerts = $true`, el diálogo se muestra y `SaveAs` **queda bloqueado
  indefinidamente** esperando a una persona.

Por eso `Build-PivotFromCsv.ps1` construye la dinámica y **deja Excel abierto sin guardar**:
el `Ctrl+S` manual es el momento en que ese diálogo se puede contestar sin colgar nada. El
modificador `-Save` fuerza el guardado automático, solo para entornos sin etiqueta
obligatoria.

Si no querés lidiar con Excel, `Get-HourlyBreakdown.ps1` da el mismo resultado en CSV.

## scripts/PivotCorreosPorHora.bas

Macro de Excel (VBA) que toma el CSV generado por el script anterior y arma una tabla
dinámica con el conteo de correos por hora del día, en formato 24 h.

### Uso

1. Abrir el CSV en Excel.
2. `Alt+F11` → **Insertar** → **Módulo** → pegar el contenido del `.bas`.
3. Volver a Excel, `Alt+F8` → `CrearPivotCorreosPorHora` → **Ejecutar**.

Agrega dos columnas auxiliares (`Hora` con el número 0-23 y `Rango horario` con
`"10:00 - 10:59"`) y crea la hoja **Correos por hora** con la dinámica y un gráfico de
barras. Si el CSV incluye la columna `Folder`, la abre como columnas de la dinámica para
comparar AWG contra Wakefern.

### Extracción de la hora

Cubre los dos casos posibles según cómo Excel haya interpretado la columna `Received`:

- **Fecha/hora real** (`VarType = vbDate` o número de serie) → `Hour()` directo.
- **Texto** (`"7/23/2026 10:16"`) → se corta después del primer espacio y se lee la parte
  anterior a los dos puntos; si VBA puede parsear ese fragmento (incluido AM/PM), se usa
  `Hour(CDate(...))`.

Las filas cuya hora no se puede determinar no se descartan en silencio: se agrupan como
`(sin hora)` y la macro informa cuántas fueron al terminar.
