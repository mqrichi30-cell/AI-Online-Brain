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
