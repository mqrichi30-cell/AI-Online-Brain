# WAKEFERN APPT Flow — Lecciones Aprendidas

## Contactos del flow (según ApptsWakefern.bas)
```
GROCERY_TO: Steve.Salotti@wakefern.com; Grocery_Special_PO_Group@wakefern.com;
            Mark.Kielczynski@wakefern.com; anne.mucchiello@wakefern.com; Al.DAgostino@wakefern.com

HABA_TO: GMScheduling@wakefern.com; sharper.d@pg.com; Irma.Lang@wakefern.com

CC_DEFAULT: jackson.vs@pg.com; rocas.cr@pg.com

OSS (RPA): nacsoshared.im@pg.com  (Subject fijo: "RPA EMAIL REQUEST TO CHANGES")
```

## Reglas de subject por email
- **Wakefern (cliente)**: `RE: {subject original del email entrante}` — mantener cadena
- **OSS/RPA**: `"RPA EMAIL REQUEST TO CHANGES"` — siempre fijo
- El subject dinámico usa `concat('RE: ', first(triggerOutputs()?['body/value'])?['subject'])`
  - ⚠️ Pendiente verificar field name exacto del subject (debug `DEBUG_TriggerItem` aún sin confirmar)

## Errores encontrados y sus causas durante desarrollo v9→v17

| Error | Causa | Fix |
|-------|-------|-----|
| "Empty Content-Type provided" | PATCH/POST sin `ContentType` | Agregar `"ContentType": "application/json"` |
| "WorkflowOperationParametersExtraParameter" | `Headers` no soportado por conector | Remover parámetro `Headers` |
| "URI path is not a valid Graph endpoint" | URI relativa (`me/messages/`) | Usar URL absoluta con `https://graph.microsoft.com/v1.0/` |
| "The action cannot reference itself" | La acción usaba `body('misma_accion')?['id']` en su propio input | Usar `triggerOutputs()` en el input de la acción que obtiene el ID |
| "OData request is not supported" | Combinación de URL absoluta + falta de ContentType | Ambos deben estar correctos a la vez |
| "ErrorItemNotFound" 404 | Resubmit sobre email movido (ID stale) | No es bug en producción; ID cambia al mover |
| Subject vacío en email enviado | `triggerOutputs()?['body/subject']` = vacío; correcto es `first(triggerOutputs()?['body/value'])?['subject']` | Ver trigger-openapi-notification-structure.md |
| DEBUG mostraba todo vacío | El Compose de debug usaba el path incorrecto, no el que se actualizó en las acciones | Siempre actualizar debug al mismo tiempo que las acciones |

## Metodología de debug recomendada
1. Agregar un único Compose `DEBUG_TriggerItem` al inicio del scope:
   ```
   @string(first(triggerOutputs()?['body/value']))
   ```
   Esto muestra la estructura completa del email con todos sus field names reales.

2. Para debugs de URI/Body en HttpRequest, agregar Compose antes de la acción que falla:
   ```
   @concat('ID: ', first(triggerOutputs()?['body/value'])?['id'],
           ' | URI: https://graph.microsoft.com/v1.0/me/messages/', first(triggerOutputs()?['body/value'])?['id'])
   ```

3. Un solo debug que consolide todo es más útil que múltiples — el usuario puede ver el output de un solo step.

## Para Move, Categorize, Mark as Read — usar Send HTTP de Outlook
Usar `operationId: "HttpRequest"` con `shared_office365`, NO el conector HTTP genérico.
Los tres siempre van en cadena: Mark Read → Categorize → Move (en ese orden).
