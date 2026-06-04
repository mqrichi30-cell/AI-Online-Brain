# Trigger OpenApiConnectionNotification — Estructura del Body

## Contexto
Trigger: "When a new email arrives (V3)" (`OnNewEmailV3`, `shared_office365`)
Tipo: `OpenApiConnectionNotification`

## CRÍTICO: el email NO está en `triggerOutputs()?['body']` directamente
Para este tipo de trigger, los datos del email están en un **array** bajo `body/value`:

```
triggerOutputs()?['body/value']  →  [ { id, subject, from, body, internetMessageId, ... } ]
```

### Path correcto para acceder a los campos del email
```
first(triggerOutputs()?['body/value'])?['id']
first(triggerOutputs()?['body/value'])?['subject']
first(triggerOutputs()?['body/value'])?['from']
first(triggerOutputs()?['body/value'])?['body']
first(triggerOutputs()?['body/value'])?['internetMessageId']
```

### Paths INCORRECTOS (devuelven vacío/null)
```
triggerOutputs()?['body/id']        ← VACÍO
triggerOutputs()?['body/subject']   ← VACÍO
triggerOutputs()?['body/from']      ← VACÍO
triggerBody()?['body']              ← VACÍO
triggerOutputs()?['body/body']      ← VACÍO
```

## Cómo detectar el problema
Si en un Compose debug todos los campos salen vacíos pero el flow SÍ entra a las condiciones, es señal de que las condiciones usan el path correcto (`body/value`) pero las acciones usan el path incorrecto (`body/field`).

## Nota sobre resubmits
Power Automate resubmit replay usa los trigger outputs originales del run. Si el email fue movido entre el run original y el resubmit, el `id` retornado por `first(triggerOutputs()?['body/value'])?['id']` será el ID antiguo → Graph devolverá 404 `ErrorItemNotFound`. Esto es un artefacto de testing, no un bug del flow en producción.

## Comparación: trigger body vs condición válida en el flow original
El flow original ya usaba el path correcto en condiciones:
```
first(triggerOutputs()?['body/value'])?['subject']
```
Pero las acciones (email send, HttpRequest) usaban el path incorrecto. Consistencia es clave.
