# Office 365 Connector — HttpRequest Limitations

## Recursos permitidos
El conector `HttpRequest` (operationId: `HttpRequest`, shared_office365) solo acepta estas rutas:
- **Recursos**: `me`, `users`
- **Objetos**: `messages`, `mailFolders`, `events`, `calendar`, `calendars`, `outlook`, `inferenceClassification`

## Parámetros soportados
Solo acepta: `Uri`, `Method`, `Body`, `ContentType`
- **NO soporta `Headers`** — cualquier intento de pasar headers custom (ej. `Prefer: IdType="ImmutableId"`) falla con `WorkflowOperationParametersExtraParameter`
- **URI debe ser absoluta**: `https://graph.microsoft.com/v1.0/me/messages/{id}` — las rutas relativas (`me/messages/{id}`) fallan con "URI path is not valid"
- **ContentType requerido en PATCH/POST**: sin él falla con "Empty Content-Type provided" o "OData request is not supported"

## Immutable ID — no es posible desde este conector
Para obtener el Immutable ID de Graph se necesita el header `Prefer: IdType="ImmutableId"` — imposible con este conector.
- `$filter=internetMessageId eq '...'` no funciona: `internetMessageId` no es propiedad filterable en Graph
- `$search` requiere permisos adicionales
- **Alternativa real**: usar el conector HTTP premium de Azure (requiere licencia)
- **Para operaciones inmediatas** (mark as read, categorize, move — dentro del mismo run): el `id` del trigger es suficiente; el ID solo cambia si el email es movido entre runs

## Ejemplo correcto de acción HttpRequest
```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "Uri": "https://graph.microsoft.com/v1.0/me/messages/{id}",
      "Method": "PATCH",
      "Body": "{\"isRead\": true}",
      "ContentType": "application/json"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
      "connectionName": "shared_office365",
      "operationId": "HttpRequest"
    },
    "authentication": "@parameters('$authentication')"
  }
}
```
